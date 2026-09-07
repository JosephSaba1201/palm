#!/usr/bin/env python3
"""
Convierte una "SOLICITUD SEMANAL DE PAGOS" (PDF) de PALM Diamante en un Excel
semanal con la misma forma que el maestro $/PALM 2026 SEM 33.xlsx (hoja
"LISTA DE PAGOS"): una fila por pago, con columnas de PARTIDA (torre) y
SUBPARTIDA (partida de obra), clasificadas por:
  1) palabras clave en el concepto (ej. "TORRE 2", "CIMENTACION"),
  2) precedente histórico del mismo proveedor (y proveedor+partida) en el
     maestro S33.

Las filas cuya clasificación quedó ambigua (sin keyword ni precedente
dominante) se marcan en amarillo para revisión rápida.

Uso:
    python3 pdf_pagos_a_excel.py "/ruta/al/PALM 2026 SEM NN.pdf" ["/ruta/al/maestro.xlsx"]

Genera: mismo folder que el PDF, "PALM 2026 SEM NN - Pagos.xlsx"
"""
import sys
import re
import json
import datetime
from pathlib import Path
from collections import defaultdict, Counter

import pdfplumber
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

SECTION_ORDER = [
    "OBRA", "FEES", "SOFT", "INTERESES",
    "PUBLICIDAD EN VENTAS", "COMISIONES POR VENTA", "DEVOLUCION POR CANCELACION",
]

CATOF_LABEL = {
    "TORRE 1": "TORRE 1", "TORRE 2": "TORRE 2", "TORRE 3": "TORRE 3",
    "GENERAL": "GENERAL", "AMENIDADES": "AMENIDADES", "SEGURIDAD": "SEGURIDAD",
    "TORRES": "TORRES", "TORRE 1 Y 2": "TORRE 1 Y 2",
    "TORRES Y AMENIDADES": "TORRES Y AMENIDADES", "VENTAS": "VENTAS",
    "MANTENIMIENTO": "MANTENIMIENTO", "MOVIMIENTO DE TIERRAS": "MOVIMIENTO DE TIERRAS",
    "CANCELACION": "CANCELACION", "SOFT": "SOFT", "FEES": "FEES",
}

KEYMAP_PARTIDA = [
    (r"TORRE\s*1\s*Y\s*2", "TORRE 1 Y 2"),
    (r"TORRES\s*Y\s*AMENIDADES", "TORRES Y AMENIDADES"),
    (r"TORRE\s*2A|TORRE\s*2B|TORRE\s*2\b|\bT2-", "TORRE 2"),
    (r"TORRE\s*1\b|\bT1-", "TORRE 1"),
    (r"TORRE\s*3\b|\bT3-", "TORRE 3"),
    (r"AMENIDADES", "AMENIDADES"),
    (r"SEGURIDAD", "SEGURIDAD"),
]

SECTION_TO_PARTIDA_FALLBACK = {
    "FEES": "FEES", "SOFT": "SOFT", "INTERESES": "SOFT",
    "PUBLICIDAD EN VENTAS": "VENTAS", "COMISIONES POR VENTA": "VENTAS",
    "DEVOLUCION POR CANCELACION": "CANCELACION",
}

KEYMAP_SUBPARTIDA = [
    (r"TOPOGRAFIA", "TOPOGRAFIA"),
    (r"CIMENTACION|PILAS\b", "CIMENTACION"),
    (r"ESTRUCTURA", "ESTRUCTURA"),
    (r"ALBA[ÑN]ILER", "ALBAÑILERIA"),
    (r"IMPERMEABILIZ", "IMPERMEABILIZACION"),
    (r"ACABADOS", "ACABADOS"),
    (r"INSTALACION", "INSTALACIONES"),
    (r"CANCELER", "CANCELERIA"),
    (r"HERRERIA", "HERRERIA"),
    (r"ELEVADOR", "ELEVADORES"),
    (r"LIMPIEZ", "LIMPIEZAS"),
    (r"DEMOLICION", "DEMOLICION"),
    (r"MOVIMIENTO DE TIERRAS|AGREGADOS|RETROEXCAVADORA|ESCOMBRO", "MOVIMIENTO DE TIERRAS"),
    (r"JARDINER", "JARDINERIA"),
    (r"CARPINTER", "CARPINTERIA"),
    (r"COCINA", "COCINAS"),
    (r"MANTENIMIENTO", "MANTENIMIENTO"),
    (r"LABORATORIO", "LABORATORIO"),
]


def clean_amount(s):
    if not s:
        return None
    s = s.replace("\n", " ").strip()
    s = re.sub(r"[A-Za-z]+", "", s)
    s = s.replace("$", "").replace(" ", "").replace(",", "").strip()
    if not s or s == "-":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def norm_name(s):
    return re.sub(r"\s+", " ", (s or "").strip().upper())


def is_amount_label_only(s):
    return bool(s) and not re.search(r"\d", s)


def extract_meta(text, pdf_filename=""):
    meta = {}
    m = re.search(r"SEM(?:ANA)?\s*(\d+)", pdf_filename, re.I)
    if m:
        meta["semana"] = m.group(1)
    m = re.search(r"A[ÑN]O\s*(\d{4})\s*/\s*SEMANA\s*(\d+)", text)
    if m:
        meta["semana_obra"] = m.group(2)
        meta.setdefault("semana", m.group(2))
    m = re.search(r"(\d{2}/\d{2}/\d{4})", text)
    if m:
        meta["fecha_pago"] = m.group(1)
    m = re.search(r"DEL\s+.*?\d{4}", text)
    if m:
        meta["periodo_obra"] = m.group(0)
    m = re.search(r"T\.?C\.?\s*USD\s*\$?\s*([\d][\d\s.,]*\d|\d)", text)
    if m:
        meta["tc_usd"] = clean_amount(m.group(1))
    return meta


def parse_pdf_rows(pdf_path: Path):
    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join(p.extract_text() or "" for p in pdf.pages)
        tables = []
        for page in pdf.pages:
            tables += page.extract_tables()

    section = None
    rows = []
    for table in tables:
        for r in table:
            r = (r + [None] * 8)[:8]
            n, benef, concepto, importe_pagar, factura, mn, usd, obs = r
            benef_c = (benef or "").strip()
            concepto_c = (concepto or "").strip()
            # Los encabezados de sección a veces vienen mal alineados por pdfplumber
            # (p.ej. la fila ['Columna1','SOFT','Columna2',...] en vez de una fila
            # limpia con 'SOFT' solo) -- se detectan buscando el nombre de sección
            # en cualquier celda de una fila que no trae un concepto real.
            cells_upper = [(c or "").strip().upper() for c in r]
            sec_match = next((c for c in cells_upper if c in SECTION_ORDER), None)
            if sec_match and (not concepto_c or concepto_c.upper().startswith("COLUMNA")):
                section = sec_match
                continue
            if not benef_c or benef_c.upper() == "TOTAL":
                continue
            mn_val = None if is_amount_label_only(mn) else clean_amount(mn)
            usd_val = None if is_amount_label_only(usd) else clean_amount(usd)
            if mn_val is None and usd_val is None:
                continue
            rows.append({
                "section": section or "OBRA",
                "n": n,
                "benef": benef_c.replace("\n", " "),
                "concepto": concepto_c.replace("\n", " "),
                "importe_pagar": clean_amount(importe_pagar),
                "factura": (factura or "").replace("\n", " ") if factura else "",
                "autorizado_mn": mn_val,
                "autorizado_usd": usd_val,
                "obs": (obs or "").replace("\n", " ") if obs else "",
            })
    return full_text, rows


CODE_RE = re.compile(r"^(\d+)([A-Z])$")


def _code_sort_key(code):
    m = CODE_RE.match(code)
    if not m:
        return None
    return (m.group(2), int(m.group(1)))  # letra de año, número de semana


def next_semana_code(master_xlsx: Path):
    """El código de SEMANA DE AÑO que le toca al siguiente lote de pagos NO
    sale de una tabla de calendario ni del nombre del PDF -- es,
    simplemente, el que sigue después del último código real que ya está
    cargado en la hoja 'LISTA DE PAGOS' del maestro (avance secuencial:
    si el maestro va pagado hasta la semana 30E, el siguiente lote es 31E,
    sin importar en qué fecha calendario caiga ni qué diga el PDF)."""
    if not master_xlsx or not master_xlsx.exists():
        return None, None, "sin maestro para determinar la última semana cargada"
    wb = openpyxl.load_workbook(master_xlsx, read_only=True, data_only=True)
    ws = wb["LISTA DE PAGOS"]
    codes = set()
    last_fecha = None
    for r in ws.iter_rows(min_row=6, values_only=True):
        if not isinstance(r[1], (datetime.date, datetime.datetime)):
            continue
        code = str(r[0]).strip().upper() if r[0] else ""
        if CODE_RE.match(code):
            codes.add(code)
            d = r[1].date()
            if last_fecha is None or d > last_fecha:
                last_fecha = d
    if not codes:
        return None, None, "no se encontraron códigos SEMANA DE AÑO (formato NNE) en el maestro"
    last_code = max(codes, key=_code_sort_key)
    n, letra = _code_sort_key(last_code)[1], _code_sort_key(last_code)[0]
    next_code = f"{n + 1}{letra}"
    return next_code, last_code, f"último cargado en el maestro: {last_code} (fecha máxima {last_fecha}) -> siguiente: {next_code}"


def load_history(master_xlsx: Path):
    """(benef -> Counter(partida)), ((benef,partida) -> Counter(subpartida))"""
    benef_partida = defaultdict(Counter)
    bp_sub = defaultdict(Counter)
    if not master_xlsx or not master_xlsx.exists():
        return benef_partida, bp_sub
    wb = openpyxl.load_workbook(master_xlsx, read_only=True, data_only=True)
    ws = wb["LISTA DE PAGOS"]
    for r in ws.iter_rows(min_row=6, values_only=True):
        if not isinstance(r[1], (datetime.date, datetime.datetime)):
            continue
        benef = norm_name(r[4])
        partida = (r[14] or "").strip().upper()
        sub = (r[15] or "").strip().upper()
        if benef and partida:
            benef_partida[benef][partida] += 1
        if benef and partida and sub:
            bp_sub[(benef, partida)][sub] += 1
    return benef_partida, bp_sub


def classify_partida(row, benef_hist):
    text = (row["concepto"] + " " + row["obs"]).upper()
    for pat, partida in KEYMAP_PARTIDA:
        if re.search(pat, text):
            return partida, "keyword", True
    if row["section"] in SECTION_TO_PARTIDA_FALLBACK and row["section"] != "OBRA":
        return SECTION_TO_PARTIDA_FALLBACK[row["section"]], "section", True
    h = benef_hist.get(norm_name(row["benef"]))
    if h:
        top, total = h.most_common(1)[0][0], sum(h.values())
        conf = h[top] / total
        if conf >= 0.55:
            return top, f"historial({h[top]}/{total})", True
        return top, f"AMBIGUO historial:{h.most_common()}", False
    return "GENERAL", "sin historial (default GENERAL)", False


def classify_subpartida(row, partida, bp_sub_hist):
    text = (row["concepto"] + " " + row["obs"]).upper()
    for pat, sub in KEYMAP_SUBPARTIDA:
        if re.search(pat, text):
            return sub, "keyword", True
    if partida in ("SOFT", "FEES"):
        return "SOFT", "partida", True
    if partida == "VENTAS":
        return "VENTAS", "partida", True
    h = bp_sub_hist.get((norm_name(row["benef"]), partida))
    if h:
        top, total = h.most_common(1)[0][0], sum(h.values())
        conf = h[top] / total
        if conf >= 0.55:
            return top, f"historial({h[top]}/{total})", True
        return top, f"AMBIGUO historial:{h.most_common()}", False
    if "MATERIALES SUMINISTRADOS" in text:
        return "MATERIALES", "keyword", True
    return "GENERAL", "sin historial (default GENERAL)", False


def hard_soft_flag(section, partida, concepto):
    if section in ("SOFT", "INTERESES", "FEES"):
        return "SOFT"
    if section == "COMISIONES POR VENTA":
        return "COMISION VENTAS"
    if section == "PUBLICIDAD EN VENTAS":
        return "VENTAS"
    if section == "DEVOLUCION POR CANCELACION":
        return "DEVOLUCION"
    if "MATERIALES SUMINISTRADO" in concepto.upper():
        return "HARD MAT"
    return "HARD"


def build_excel(pdf_path: Path, master_xlsx: Path, out_path: Path, semana_override: str = None):
    full_text, rows = parse_pdf_rows(pdf_path)
    meta = extract_meta(full_text, pdf_path.name)
    benef_hist, bp_sub_hist = load_history(master_xlsx)

    fecha_pago = None
    if meta.get("fecha_pago"):
        d, m, y = meta["fecha_pago"].split("/")
        fecha_pago = datetime.date(int(y), int(m), int(d))

    # El número de SEMANA DE AÑO NO se infiere solo (ni del nombre del PDF,
    # ni de una tabla de calendario): los pagos se pueden cargar fuera de
    # orden (ej. Joseph avisó que va a cargar las semanas 31 y 32 DESPUÉS
    # de procesar la 33). Por default se sugiere "el siguiente después del
    # último cargado en el maestro", pero SIEMPRE hay que confirmar con
    # Joseph qué semana es cada lote -- `semana_override` es esa confirmación.
    suggested_code, last_code, reason = next_semana_code(master_xlsx)
    if semana_override:
        meta["semana_maestro"] = semana_override.upper()
        meta["semana_lookup_reason"] = f"confirmado manualmente (sugerencia automática habría sido: {suggested_code})"
    elif suggested_code:
        meta["semana_maestro"] = suggested_code
        meta["semana_lookup_reason"] = reason
    meta["semana_periodo_maestro"] = None

    tc = meta.get("tc_usd")

    enriched = []
    ambiguous_count = 0
    for row in rows:
        partida, p_reason, p_ok = classify_partida(row, benef_hist)
        sub, s_reason, s_ok = classify_subpartida(row, partida, bp_sub_hist)
        importe = row["autorizado_mn"] if row["autorizado_mn"] else (
            (row["autorizado_usd"] or 0) * tc if tc and row["autorizado_usd"] else None)
        sin_iva = importe / 1.16 if importe else None
        iva = importe - sin_iva if importe else None
        hs_flag = hard_soft_flag(row["section"], partida, row["concepto"])
        flagged = not (p_ok and s_ok)
        if flagged:
            ambiguous_count += 1
        enriched.append({
            **row, "partida": partida, "p_reason": p_reason, "p_ok": p_ok,
            "subpartida": sub, "s_reason": s_reason, "s_ok": s_ok,
            "importe": importe, "sin_iva": sin_iva, "iva": iva,
            "hs_flag": hs_flag, "flagged": flagged,
        })

    semana_display = meta.get("semana_maestro") or meta.get("semana", "XX")

    wb = Workbook()
    ws = wb.active
    ws.title = f"SEM {semana_display}"[:31]

    headers = ["SEMANA DE AÑO", "FECHA", "HARD, SOFT, FUERA DE OBRA", "N°", "BENEFICIARIO",
               "PAGADO SIN IVA", "IVA DEL 16%", "IMPORTE", "PAGADO EN DOLARES", "TIPO DE CAMBIO",
               "SUBPARTIDA", "FACTURA / REF.", "CONCEPTO", "PARTIDA", "OBSERVACIONES", "REVISAR"]

    bold = Font(bold=True)
    header_fill = PatternFill("solid", fgColor="1F2937")
    header_font = Font(bold=True, color="FFFFFF")
    flag_fill = PatternFill("solid", fgColor="FFF3B0")
    thin = Side(style="thin", color="D1D5DB")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    money_fmt = "#,##0.00"

    ws["A1"] = "PALM DIAMANTE — Registro de Pagos"
    ws["A1"].font = Font(bold=True, size=13)
    ws.merge_cells("A1:P1")
    obra_txt = meta.get("periodo_obra") or ""
    ws["A2"] = (f"Semana {semana_display}  (buscador de semanas, según fecha de pago)"
                f"  ·  Fecha de pago: {meta.get('fecha_pago','?')}"
                + (f"  ·  Obra facturada: {obra_txt}" if obra_txt else "")
                + (f"  ·  T.C. USD: {tc}" if tc else ""))
    ws["A2"].font = Font(italic=True)
    ws.merge_cells("A2:P2")

    header_row = 4
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border

    r = header_row + 1
    for row in enriched:
        vals = [
            semana_display, fecha_pago, row["hs_flag"], row["n"], row["benef"],
            row["sin_iva"], row["iva"], row["importe"], row["autorizado_usd"], tc if row["autorizado_usd"] else None,
            row["subpartida"], row["factura"], row["concepto"], row["partida"], row["obs"],
            "" if not row["flagged"] else f"partida:{row['p_reason']} | sub:{row['s_reason']}",
        ]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = border
            if c in (6, 7, 8, 9):
                cell.number_format = money_fmt
            if c == 2 and v:
                cell.number_format = "dd/mm/yyyy"
            if row["flagged"]:
                cell.fill = flag_fill
        r += 1

    total_row = r + 1
    ws.cell(row=total_row, column=5, value="TOTAL").font = bold
    ws.cell(row=total_row, column=8, value=sum(x["importe"] or 0 for x in enriched)).font = bold
    ws.cell(row=total_row, column=8).number_format = money_fmt
    ws.cell(row=total_row, column=9, value=sum(x["autorizado_usd"] or 0 for x in enriched)).font = bold
    ws.cell(row=total_row, column=9).number_format = money_fmt

    widths = [10, 12, 16, 6, 32, 14, 12, 14, 12, 10, 16, 16, 42, 14, 26, 40]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"

    wb.save(out_path)
    return meta, enriched, ambiguous_count


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python3 pdf_pagos_a_excel.py <ruta_pdf> [<ruta_maestro_xlsx>] [<semana_confirmada, ej. 33E>]")
        sys.exit(1)
    pdf_path = Path(sys.argv[1])
    master_xlsx = Path(sys.argv[2]) if len(sys.argv) > 2 else None
    semana_override = sys.argv[3] if len(sys.argv) > 3 else None
    tmp_path = pdf_path.parent / f"~tmp_{pdf_path.stem}.xlsx"
    meta, enriched, amb = build_excel(pdf_path, master_xlsx, tmp_path, semana_override)
    semana_display = meta.get("semana_maestro") or meta.get("semana", "XX")
    out_path = pdf_path.parent / f"PALM 2026 SEM {semana_display} - Pagos.xlsx"
    tmp_path.replace(out_path)
    total = sum(x["importe"] or 0 for x in enriched)
    total_usd = sum(x["autorizado_usd"] or 0 for x in enriched)
    print(f"OK -> {out_path}")
    print(f"Semana (maestro): {semana_display}  [{meta.get('semana_lookup_reason')}]")
    print(f"Semana según nombre del PDF: {meta.get('semana')}  ·  Fecha pago: {meta.get('fecha_pago')}")
    if meta.get("semana_periodo_maestro"):
        print(f"Periodo de obra (buscador de semanas): {meta['semana_periodo_maestro']}")
    print(f"Filas: {len(enriched)}  Total MN: {total:,.2f}  Total USD: {total_usd:,.2f}")
    print(f"Filas marcadas para revisar (partida/subpartida ambigua): {amb}")
