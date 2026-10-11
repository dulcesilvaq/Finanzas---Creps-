"""
Genera los archivos .txt para subir al BCP (crep_s_DDMM.txt / crep_d_DDMM.txt)
a partir del Excel con las hojas "Código libre soles" y "Código libre dólares".

Estructura (verificada contra crep_s_1209.txt y crep_d_1209.txt):

  CABECERA (1 sola línea, SIEMPRE la primera, 200 caracteres)
    CC | cuenta(11) | C | nombre(40, relleno con espacios) | fecha AAAAMMDD(8)
       | n° de registros(9) | monto total en centavos(15) | A | 110 espacios | 000

  DETALLE (una línea por fila del Excel, 153 caracteres) = columna H del Excel

  Separador de líneas: CRLF, sin salto de línea después de la última (igual que los ejemplos).
"""
import io
import re
from datetime import datetime, timedelta, timezone

from openpyxl import load_workbook

import crep

HOJAS = {
    "S": crep.HOJA_COD_S,
    "D": crep.HOJA_COD_D,
}

# cuenta = convenio sin el prefijo "DD"; nombre tal cual aparece en los archivos de ejemplo
CONFIG = {
    "S": {"cuenta": "21501863341", "nombre": "BSGRUPOLIMASOLES",
          "conv": crep.CONV_SOLES, "fproc": crep.FPROC_SOLES},
    "D": {"cuenta": "21511870934", "nombre": "BSGINSTITUTELIMADOLARES",
          "conv": crep.CONV_DOLARES, "fproc": crep.FPROC_DOLARES},
}

LARGO_DETALLE = 153
LIMA = timezone(timedelta(hours=-5))      # Perú no tiene horario de verano


def hoy_lima() -> datetime:
    return datetime.now(LIMA)


# ---------------------------------------------------------------- lectura del Excel
def leer_codigos(contenido: bytes, tipo: str):
    """
    Lee A..H de la hoja de códigos. Devuelve (filas, avisos).
    Cada fila = dict con la línea final 'linea' y datos para validar.
    Si H no tiene valor guardado (fórmula sin calcular), se calcula con la misma
    lógica de la fórmula a partir de A..G.
    """
    cfg = CONFIG[tipo]
    nombre_hoja = HOJAS[tipo]
    avisos = []
    wb = load_workbook(io.BytesIO(contenido), read_only=True, data_only=True)
    if nombre_hoja not in wb.sheetnames:
        raise ValueError(f'No encuentro la hoja "{nombre_hoja}" en el Excel.')
    ws = wb[nombre_hoja]

    filas, calculadas, manuales = [], 0, 0
    for r, row in enumerate(ws.iter_rows(min_row=2, max_col=8, values_only=True), start=2):
        row = list(row) + [None] * (8 - len(row))
        a, b, c, d, e, f, g, h = row
        if all(v is None or str(v).strip() == "" for v in row):
            continue                                   # fila vacía
        a_s, c_s = crep.a_texto(a), crep.a_texto(c)
        esperado = None
        try:
            esperado = crep.crep_valor(a_s, crep.a_texto(b), c_s, crep.a_texto(d), crep.a_texto(e),
                                       crep.a_texto(f), crep.num(g), cfg["conv"], cfg["fproc"])
        except Exception:
            pass
        if h is None or str(h) == "":
            if esperado is None:
                avisos.append(f"Fila {r}: sin código en H y no se pudo calcular.")
                continue
            linea = esperado
            calculadas += 1
        else:
            linea = str(h)
            if esperado is not None and linea != esperado:
                manuales += 1
        filas.append({"fila": r, "linea": linea})

    if calculadas:
        avisos.append(f"{calculadas} filas no tenían el código H guardado y se calcularon con la misma fórmula.")
    return filas, avisos, manuales


# ---------------------------------------------------------------- validación
def validar_linea(linea: str, cfg) -> list:
    """Devuelve la lista de problemas de una línea de detalle (vacía = OK)."""
    p = []
    if len(linea) != LARGO_DETALLE:
        p.append(f"largo {len(linea)} (debe ser {LARGO_DETALLE})")
        return p                                       # el resto depende de las posiciones
    if not linea.isascii() or not linea.isprintable():
        p.append("tiene caracteres no permitidos (tildes, tabulaciones o saltos de línea)")
    if not linea.startswith(cfg["conv"]):
        p.append(f"no empieza con el convenio {cfg['conv']}")
    if linea[78:97] != " " * 19:
        p.append("espacios fijos (posiciones 79-97) alterados")
    if not re.fullmatch(r"1\d\d0\dXXXXXX", linea[67:78]):
        p.append(f"código de cuota raro: '{linea[67:78]}'")
    if linea[97:105] != cfg["fproc"]:
        p.append(f"fecha de proceso {linea[97:105]} distinta de {cfg['fproc']}")
    try:
        datetime.strptime(linea[105:113], "%Y%m%d")
    except ValueError:
        p.append(f"fecha de vencimiento inválida '{linea[105:113]}'")
    if not linea[113:128].isdigit():
        p.append("monto con caracteres no numéricos")
    if linea[128:152] != "0" * 24 or linea[152] != "A":
        p.append("cierre (24 ceros + 'A') alterado")
    return p


# ---------------------------------------------------------------- armado del txt
def armar_cabecera(cfg, fecha: datetime, n: int, total_centavos: int) -> str:
    cab = ("CC" + cfg["cuenta"] + "C" + cfg["nombre"].ljust(40)
           + fecha.strftime("%Y%m%d") + str(n).zfill(9) + str(total_centavos).zfill(15)
           + "A" + " " * 110 + "000")
    assert len(cab) == 200, len(cab)
    return cab


def generar_txt(contenido: bytes, tipo: str, fecha: datetime):
    """tipo: 'S' o 'D'. Devuelve (bytes_txt, resumen dict, avisos)."""
    cfg = CONFIG[tipo]
    filas, avisos, manuales = leer_codigos(contenido, tipo)

    problemas = []
    for f in filas:
        for msg in validar_linea(f["linea"], cfg):
            problemas.append(f"Fila {f['fila']} de '{HOJAS[tipo]}': {msg}")

    lineas = [f["linea"] for f in filas]
    # mismo alumno + misma cuota repetidos = riesgo de cobro doble
    claves = [(l[13:27], l[67:78]) for l in lineas if len(l) == LARGO_DETALLE]
    repetidas = len(claves) - len(set(claves))
    if repetidas:
        avisos.append(f"{repetidas} líneas repiten matrícula + cuota (posible doble cobro). Revísalas.")
    if len(lineas) != len(set(lineas)):
        avisos.append(f"Hay {len(lineas) - len(set(lineas))} líneas idénticas repetidas.")

    total = sum(int(l[113:128]) for l in lineas if len(l) == LARGO_DETALLE and l[113:128].isdigit())
    cabecera = armar_cabecera(cfg, fecha, len(lineas), total)
    texto = "\r\n".join([cabecera] + lineas)           # CRLF, sin salto final (como los ejemplos)
    nombre_archivo = f"crep_{tipo.lower()}_{fecha.strftime('%d%m')}.txt"
    resumen = {
        "archivo": nombre_archivo, "registros": len(lineas), "total_centavos": total,
        "total": total / 100, "fecha": fecha.strftime("%d/%m/%Y"),
        "modificadas_a_mano": manuales, "problemas": len(problemas),
    }
    return texto.encode("ascii"), resumen, avisos, problemas
