"""Convierte los documentos de la carpeta 'input' (PDF, DOC, DOCX, PPT, PPTX)
a archivos TXT en la carpeta 'output'."""

import glob
import re
import struct
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox

import olefile
from docx import Document
from pptx import Presentation
from pypdf import PdfReader

# Como .exe (PyInstaller) se usa la carpeta del ejecutable, no la temporal donde se descomprime.
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
TAMANO_MAXIMO = 150 * 1024  # bytes por archivo TXT (150 KB, como lo muestra Windows)


# ---------------------------------------------------------------- PDF

def extraer_pdf(ruta: Path) -> str:
    reader = PdfReader(ruta)
    return "\n\n".join((page.extract_text() or "") for page in reader.pages)


# ---------------------------------------------------------------- DOCX

def extraer_docx(ruta: Path) -> str:
    doc = Document(ruta)
    lineas = []
    # Recorre el cuerpo en orden para respetar la posición de las tablas.
    for bloque in doc.element.body.iterchildren():
        tag = bloque.tag.rsplit("}", 1)[-1]
        if tag == "p":
            lineas.append("".join(t.text or "" for t in bloque.iter() if t.tag.endswith("}t")))
        elif tag == "tbl":
            for fila in bloque.iter():
                if not fila.tag.endswith("}tr"):
                    continue
                celdas = []
                for celda in fila.iterchildren():
                    if celda.tag.endswith("}tc"):
                        celdas.append(" ".join(
                            t.text or "" for t in celda.iter() if t.tag.endswith("}t")).strip())
                lineas.append("\t".join(celdas))
    return "\n".join(lineas)


# ---------------------------------------------------------------- PPTX

def _texto_shape(shape) -> list[str]:
    lineas = []
    if shape.shape_type == 6:  # grupo
        for sub in shape.shapes:
            lineas += _texto_shape(sub)
    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
        lineas += [p.text for p in shape.text_frame.paragraphs]
    if getattr(shape, "has_table", False) and shape.has_table:
        for fila in shape.table.rows:
            lineas.append("\t".join(c.text for c in fila.cells))
    return lineas


def extraer_pptx(ruta: Path) -> str:
    prs = Presentation(ruta)
    partes = []
    for i, slide in enumerate(prs.slides, 1):
        lineas = [f"--- Diapositiva {i} ---"]
        for shape in slide.shapes:
            lineas += _texto_shape(shape)
        if slide.has_notes_slide:
            notas = slide.notes_slide.notes_text_frame.text.strip()
            if notas:
                lineas += ["[Notas]", notas]
        partes.append("\n".join(lineas))
    return "\n\n".join(partes)


# ---------------------------------------------------------------- DOC (Word 97-2003)

def _limpiar_texto_word(texto: str) -> str:
    # Campos: \x13 instrucciones \x14 resultado \x15 -> se conserva solo el resultado.
    salida, pila = [], []
    for ch in texto:
        if ch == "\x13":
            pila.append(False)          # False = aún en la instrucción
        elif ch == "\x14" and pila:
            pila[-1] = True             # True = en el resultado
        elif ch == "\x15" and pila:
            pila.pop()
        elif not pila or all(pila):
            salida.append(ch)
    texto = "".join(salida)
    texto = texto.replace("\r\x07", "\n").replace("\x07", "\t")
    texto = re.sub(r"[\r\x0b\x0c\x0e]", "\n", texto)
    return re.sub(r"[\x00-\x08\x0f-\x1f]", "", texto)


def _tabla_de_piezas(clx: bytes) -> tuple[bytes, int, tuple]:
    """Recorre el Clx hasta la tabla de piezas (Pcdt). Devuelve (plc, n, cps)."""
    i = 0
    while i < len(clx) and clx[i] == 0x01:
        i += 3 + struct.unpack_from("<H", clx, i + 1)[0]
    if i >= len(clx) or clx[i] != 0x02:
        raise ValueError("estructura del documento Word no reconocida")
    lcb = struct.unpack_from("<I", clx, i + 1)[0]
    plc = clx[i + 5:i + 5 + lcb]
    n = (lcb - 4) // 12
    return plc, n, struct.unpack_from(f"<{n + 1}i", plc, 0)


def _codificacion_por_idioma(lid: int) -> str:
    primario = lid & 0x3FF
    if primario in (0x02, 0x19, 0x22, 0x23, 0x2F, 0x1A) and lid != 0x041A:
        return "cp1251"  # cirílico (búlgaro, ruso, ucraniano, bielorruso, macedonio, serbio)
    return {0x05: "cp1250", 0x0E: "cp1250", 0x15: "cp1250", 0x1B: "cp1250", 0x24: "cp1250",
            0x1A: "cp1250", 0x18: "cp1250", 0x08: "cp1253", 0x1F: "cp1254", 0x0D: "cp1255",
            0x01: "cp1256", 0x25: "cp1257", 0x26: "cp1257", 0x27: "cp1257"}.get(primario, "cp1252")


def _extraer_word95(wd: bytes) -> str:
    """Word 6.0 / 95: el texto es de 1 byte por carácter (según el idioma del documento)."""
    flags = struct.unpack_from("<H", wd, 0x0A)[0]
    if flags & 0x0100:
        raise ValueError("el documento está protegido con contraseña")
    cod = _codificacion_por_idioma(struct.unpack_from("<H", wd, 6)[0])
    ccp_text = struct.unpack_from("<i", wd, 0x34)[0]
    if not flags & 0x0004:  # documento no "complejo": texto contiguo
        fc_min = struct.unpack_from("<I", wd, 0x18)[0]
        return wd[fc_min:fc_min + ccp_text].decode(cod, errors="replace")
    fc_clx, lcb_clx = struct.unpack_from("<II", wd, 0x160)
    plc, n, cps = _tabla_de_piezas(wd[fc_clx:fc_clx + lcb_clx])
    texto = []
    for k in range(n):
        if cps[k] >= ccp_text:
            break
        largo = min(cps[k + 1], ccp_text) - cps[k]
        fc = struct.unpack_from("<I", plc, (n + 1) * 4 + k * 8 + 2)[0]
        texto.append(wd[fc:fc + largo].decode(cod, errors="replace"))
    return "".join(texto)


def extraer_doc(ruta: Path) -> str:
    with olefile.OleFileIO(str(ruta)) as ole:
        wd = ole.openstream("WordDocument").read()
        ident = struct.unpack_from("<H", wd, 0)[0]
        if ident == 0xA5DC:
            return _limpiar_texto_word(_extraer_word95(wd))
        if ident != 0xA5EC:
            raise ValueError("no es un documento Word válido")
        flags = struct.unpack_from("<H", wd, 0x0A)[0]
        if flags & 0x0100:
            raise ValueError("el documento está protegido con contraseña")
        tabla = "1Table" if flags & 0x0200 else "0Table"
        tbl = ole.openstream(tabla).read()

    # Word 97-2003: localiza los campos del FIB de forma dinámica.
    csw = struct.unpack_from("<H", wd, 32)[0]
    pos_lw = 34 + csw * 2
    cslw = struct.unpack_from("<H", wd, pos_lw)[0]
    ccp_text = struct.unpack_from("<i", wd, pos_lw + 2 + 3 * 4)[0]
    pos_fclcb = pos_lw + 2 + cslw * 4 + 2
    fc_clx, lcb_clx = struct.unpack_from("<II", wd, pos_fclcb + 33 * 8)
    plc, n, cps = _tabla_de_piezas(tbl[fc_clx:fc_clx + lcb_clx])

    texto = []
    for k in range(n):
        if cps[k] >= ccp_text:
            break
        largo = min(cps[k + 1], ccp_text) - cps[k]
        fc = struct.unpack_from("<I", plc, (n + 1) * 4 + k * 8 + 2)[0]
        if fc & 0x40000000:  # texto comprimido (1 byte por carácter, cp1252)
            ini = (fc & 0x3FFFFFFF) // 2
            texto.append(wd[ini:ini + largo].decode("cp1252", errors="replace"))
        else:                # UTF-16
            ini = fc & 0x3FFFFFFF
            texto.append(wd[ini:ini + largo * 2].decode("utf-16-le", errors="replace"))
    return _limpiar_texto_word("".join(texto))


# ---------------------------------------------------------------- PPT (PowerPoint 97-2003)

def extraer_ppt(ruta: Path) -> str:
    with olefile.OleFileIO(str(ruta)) as ole:
        if ole.exists("EncryptedSummary"):
            raise ValueError("la presentación está protegida con contraseña")
        data = ole.openstream("PowerPoint Document").read()

    DOCUMENTO, SLWT, NOTAS, NOTES_ATOM = 0x03E8, 0x0FF0, 0x03F0, 0x03F1
    MAESTRAS = {0x03F8, 0x0FC9}  # MainMaster y HandoutContainer

    def hijos(ini: int, fin: int):
        i = ini
        while i + 8 <= fin:
            ver_inst, tipo, largo = struct.unpack_from("<HHI", data, i)
            final = min(i + 8 + largo, fin)
            yield ver_inst, tipo, i + 8, final
            i = final

    def es_maestra_de_notas(ini: int, fin: int) -> bool:
        for _, tipo, a, _ in hijos(ini, fin):
            if tipo == NOTES_ATOM:
                # Las notas reales apuntan a una diapositiva (id 0x100..0x7FFFFFFF).
                id_diapositiva = struct.unpack_from("<I", data, a)[0]
                return not 0x100 <= id_diapositiva <= 0x7FFFFFFF
        return False

    # Con "guardado rápido" quedan versiones viejas del documento; se usa la última.
    documentos = [(a, b) for _, tipo, a, b in hijos(0, len(data)) if tipo == DOCUMENTO]
    ultimo_documento = documentos[-1][0] if documentos else None

    bloques, vistos = [], set()

    def recorrer(ini: int, fin: int) -> None:
        for ver_inst, tipo, a, b in hijos(ini, fin):
            if (ver_inst & 0x0F) == 0x0F:      # contenedor
                if (tipo in MAESTRAS
                        or (tipo == DOCUMENTO and a != ultimo_documento)
                        or (tipo == SLWT and ver_inst >> 4 == 1)
                        or (tipo == NOTAS and es_maestra_de_notas(a, b))):
                    continue                   # texto de plantilla o versión vieja
                recorrer(a, b)
            elif tipo in (0x0FA0, 0x0FA8):     # TextCharsAtom / TextBytesAtom
                cuerpo = data[a:b]
                texto = (cuerpo.decode("utf-16-le", errors="replace") if tipo == 0x0FA0
                         else cuerpo.decode("cp1252", errors="replace"))
                texto = re.sub(r"[\r\x0b]", "\n", texto).strip("\x00 \n\t")
                # La misma frase suele guardarse varias veces; se evita duplicarla.
                if texto and texto != "*" and texto not in vistos:
                    vistos.add(texto)
                    bloques.append(texto)

    # Primero el texto de las diapositivas (en el documento), luego cuadros de texto y notas.
    if documentos:
        recorrer(*documentos[-1])
    recorrer(0, len(data))
    return "\n\n".join(bloques)


# ---------------------------------------------------------------- Proceso general

EXTRACTORES = {
    ".pdf": extraer_pdf,
    ".docx": extraer_docx,
    ".pptx": extraer_pptx,
    ".doc": extraer_doc,
    ".ppt": extraer_ppt,
}


def _bytes_en_disco(texto: str) -> int:
    # UTF-8, y en Windows cada "\n" se guarda como "\r\n" (un byte más).
    return len(texto.encode("utf-8")) + texto.count("\n")


def _cortar_linea(linea: str, maximo: int) -> list[str]:
    """Corta una línea demasiado larga, preferentemente en un espacio."""
    piezas = []
    while _bytes_en_disco(linea) > maximo:
        # Búsqueda binaria del prefijo más largo que cabe (sin romper caracteres).
        lo, hi = 1, len(linea)
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if _bytes_en_disco(linea[:mid]) <= maximo:
                lo = mid
            else:
                hi = mid - 1
        corte = linea.rfind(" ", 0, lo) + 1 or lo
        piezas.append(linea[:corte])
        linea = linea[corte:]
    piezas.append(linea)
    return piezas


def dividir_texto(texto: str, maximo: int = TAMANO_MAXIMO) -> list[str]:
    """Divide el texto en partes de como máximo `maximo` bytes, cortando al final de línea."""
    partes, actual, tamano = [], [], 0
    for linea in texto.splitlines(keepends=True):
        piezas = _cortar_linea(linea, maximo) if _bytes_en_disco(linea) > maximo else [linea]
        for pieza in piezas:
            b = _bytes_en_disco(pieza)
            if actual and tamano + b > maximo:
                partes.append("".join(actual))
                actual, tamano = [], 0
            actual.append(pieza)
            tamano += b
    if actual or not partes:
        partes.append("".join(actual))
    return partes


def convertir_todos() -> tuple[int, list[str]]:
    """Convierte todos los archivos soportados. Devuelve (convertidos, errores)."""
    INPUT_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)
    archivos = sorted(
        p for p in INPUT_DIR.iterdir()
        if p.is_file() and p.suffix.lower() in EXTRACTORES
        and not p.name.startswith("~$")  # archivos temporales de Office
    )
    if not archivos:
        raise FileNotFoundError(
            "No se encontraron archivos PDF, DOC, DOCX, PPT o PPTX en la carpeta 'input'.")

    # Si hay nombres repetidos (p. ej. 'informe.pdf' e 'informe.docx') se añade la extensión.
    nombres = [p.stem.lower() for p in archivos]
    convertidos, errores = 0, []
    for archivo in archivos:
        try:
            texto = EXTRACTORES[archivo.suffix.lower()](archivo)
            nombre = archivo.stem
            if nombres.count(archivo.stem.lower()) > 1:
                nombre += "_" + archivo.suffix.lower().lstrip(".")
            # Se borran las salidas anteriores de este documento para no dejar partes viejas.
            patron = re.compile(rf"{re.escape(nombre)}(_parte\d+)?\.txt", re.IGNORECASE)
            for viejo in OUTPUT_DIR.glob(f"{glob.escape(nombre)}*.txt"):
                if patron.fullmatch(viejo.name):
                    viejo.unlink()
            partes = dividir_texto(texto)
            if len(partes) == 1:
                (OUTPUT_DIR / f"{nombre}.txt").write_text(partes[0], encoding="utf-8")
            else:
                for n, parte in enumerate(partes, 1):
                    (OUTPUT_DIR / f"{nombre}_parte{n}.txt").write_text(parte, encoding="utf-8")
            convertidos += 1
        except Exception as e:
            errores.append(f"{archivo.name}: no se pudo convertir ({e})")
    return convertidos, errores


# ---------------------------------------------------------------- Interfaz

class App:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        root.title("Documentos a TXT")
        root.geometry("340x170")
        root.resizable(False, False)

        tk.Label(root, text="PDF · DOC · DOCX · PPT · PPTX", font=("Segoe UI", 9)).pack(pady=(15, 0))
        self.boton = tk.Button(root, text="Convertir a TXT", font=("Segoe UI", 12),
                               width=18, command=self.iniciar)
        self.boton.pack(pady=15)

        self.estado = tk.Label(root, text="", font=("Segoe UI", 12, "bold"))
        self.estado.pack()

    def iniciar(self) -> None:
        self.boton.config(state="disabled")
        self.estado.config(text="Procesando...", fg="black")
        threading.Thread(target=self.trabajar, daemon=True).start()

    def trabajar(self) -> None:
        try:
            _, errores = convertir_todos()
            self.root.after(0, self.terminar, errores)
        except Exception as e:
            self.root.after(0, self.terminar, [str(e)])

    def terminar(self, errores: list[str]) -> None:
        self.boton.config(state="normal")
        if not errores:
            self.estado.config(text="Finalizado", fg="green")
            messagebox.showinfo("Documentos a TXT", "Finalizado")
        else:
            self.estado.config(text="Error", fg="red")
            messagebox.showerror("Documentos a TXT", "Error\n\n" + "\n".join(errores))


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
