# -*- coding: utf-8 -*-
"""Readers for the files of one MercadoLibre download folder — this report's own.

Each reader returns a pandas DataFrame with short English column names (and,
where useful, a small dict of facts about the file). Header rows are found by
the column names they must contain, never by a fixed row number: MercadoLibre
inserts notice rows above its tables.

Roles (README): only the invoices (Facturación, Notas de crédito) give amounts.
The sales export, MercadoPago files, Full and stock files are read for SKU
lookup, status evidence and per-SKU attribution only.
"""
import datetime
import glob
import json
import os
import re
import warnings

import pandas as pd

warnings.filterwarnings("ignore", message="Workbook contains no default style")

# ------------------------------------------------------------------ small helpers

MONTHS_ES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7,
             "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11,
             "diciembre": 12}
MONTH_ABBR = {"ene": 1, "jan": 1, "feb": 2, "mar": 3, "abr": 4, "apr": 4, "may": 5, "jun": 6,
              "jul": 7, "ago": 8, "aug": 8, "sep": 9, "sept": 9, "set": 9, "oct": 10, "nov": 11,
              "dic": 12, "dec": 12}


def idstr(v):
    """An id as text: 2000018220767084.0 -> '2000018220767084'; blanks -> ''."""
    if v is None:
        return ""
    if isinstance(v, float):
        if v != v:
            return ""
        if v.is_integer():
            return str(int(v))
    if isinstance(v, int):
        return str(v)
    s = str(v).strip()
    if s.lower() in ("nan", "none", "-"):
        return ""
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def text(v):
    if v is None or (isinstance(v, float) and v != v):
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def num(v):
    """A number from a cell; '' / '-' / text -> 0.0."""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return 0.0 if v != v else float(v)
    s = str(v).strip().replace(",", "")
    if s in ("", "-", "nan"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        m = re.match(r"^-?\d+(\.\d+)?", s)
        return float(m.group(0)) if m else 0.0


def listing_digits(v):
    """'MLM5929009984' / '5929009984' / 5929009984.0 -> '5929009984'."""
    s = idstr(v)
    m = re.search(r"(\d{6,})", s)
    return m.group(1) if m else ""


def to_date(v):
    if v is None or (isinstance(v, float) and v != v):
        return None
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    if isinstance(v, pd.Timestamp):
        return None if pd.isna(v) else v.date()
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.datetime.strptime(s[:10], fmt).date()
        except ValueError:
            pass
    return spanish_date(s)


_ES_DATE = re.compile(r"(\d{1,2})\s+de\s+([a-záéíóú]+)\s+(?:de\s+)?(\d{4})", re.I)


def spanish_date(s):
    """'13 de octubre de 2026 13:35 hs.' -> date(2026, 10, 13); None if not a date."""
    m = _ES_DATE.search(str(s or ""))
    if not m:
        return None
    mon = MONTHS_ES.get(m.group(2).lower())
    if not mon:
        return None
    try:
        return datetime.date(int(m.group(3)), mon, int(m.group(1)))
    except ValueError:
        return None


def file_month(name):
    """'Reporte_Facturacion_MercadoLibre_Sept2026.xlsx' -> '2026-09' (or None)."""
    m = re.search(r"_([A-Za-z]+)(\d{4})\.xlsx$", name)
    if not m:
        return None
    mon = MONTH_ABBR.get(m.group(1).lower()) or MONTH_ABBR.get(m.group(1).lower()[:3])
    return "%s-%02d" % (m.group(2), mon) if mon else None


def header_index(raw, required, scan=40):
    """Row index of the first row (within `scan`) holding every name in `required`."""
    for i in range(min(scan, len(raw))):
        vals = set(text(v) for v in raw.iloc[i].tolist())
        if all(r in vals for r in required):
            return i
    return None


def frame_below(raw, i):
    """The table under header row i, with unique column names (dupes get .1, .2)."""
    cols, seen = [], {}
    for v in raw.iloc[i].tolist():
        name = text(v)
        if name in seen:
            seen[name] += 1
            cols.append("%s.%d" % (name, seen[name]))
        else:
            seen[name] = 0
            cols.append(name)
    df = raw.iloc[i + 1:].copy()
    df.columns = cols
    return df.reset_index(drop=True)


def read_raw(path, sheet=0):
    return pd.read_excel(path, sheet_name=sheet, header=None, dtype=object)


def updated_at(raw):
    """'Fecha de actualización' in the first rows, as text."""
    for i in range(min(5, len(raw))):
        row = raw.iloc[i].tolist()
        if text(row[0]).startswith("Fecha de actualiz"):
            for v in row[1:]:
                if isinstance(v, (datetime.datetime, pd.Timestamp)):
                    return v.strftime("%Y-%m-%d %H:%M")
                if text(v):
                    return text(v)[:16]
    return ""


def first(folder, pattern):
    hits = sorted(glob.glob(os.path.join(folder, pattern)))
    return hits[-1] if hits else None


def all_files(folder, pattern):
    return sorted(glob.glob(os.path.join(folder, pattern)))


def run_meta(folder):
    p = os.path.join(folder, "run_meta.json")
    if not os.path.exists(p):
        return {}
    try:
        with open(p, encoding="utf-8") as fh:
            return json.load(fh)
    except (ValueError, OSError):
        return {}


# ------------------------------------------------------------------ invoices

INV_COLS = {
    "N° de factura fiscal": "fiscal_no", "Fecha del cargo": "charge_date", "Número del cargo": "charge_no",
    "Detalle": "detalle", "Descontado de la operación": "deducted", "Estado del cargo": "charge_state",
    "Cargo que bonifica": "ref", "Cargo que anula": "ref", "Valor del cargo": "amount",
    "Porcentaje por categoría": "pct", "Costo fijo": "fixed_fee",
    "Subtotal sin descuento": "fee_gross", "Valor del descuento": "fee_discount",
    "Motivo del descuento": "discount_reason", "Número de venta": "sale_no", "Pago": "pago",
    "Fecha de venta": "sale_date", "Cantidad vendida": "qty", "Precio unitario": "unit_price",
    "Total de la venta": "sale_total", "Número de envío": "ship_no", "Número de paquete": "pack_no",
    "Envío a cargo del cliente": "buyer_ship", "Número de publicación": "listing",
    "Número de kit": "kit", "Título de publicación": "title", "Código ML": "inv_code",
}
INV_ID_COLS = ("fiscal_no", "charge_no", "ref", "sale_no", "pago", "ship_no", "pack_no", "listing", "kit")
INV_NUM_COLS = ("amount", "pct", "fixed_fee", "fee_gross", "fee_discount", "qty", "unit_price",
                "sale_total", "buyer_ship")


# The note MercadoLibre puts on top of a bill whose period has not closed yet:
# "Este reporte podría tener cambios hasta el cierre del período de facturación".
BILL_OPEN_RE = re.compile(r"podr[ií]a tener cambios hasta el cierre", re.I)


def bill_open(raw, header_row):
    """True when the rows above the header carry the "not closed yet" note."""
    top = raw.iloc[:max(header_row, 1)].values.ravel()
    return any(isinstance(v, str) and BILL_OPEN_RE.search(v) for v in top)


def read_invoice(path):
    """One Facturación or Notas de crédito file -> (rows, facts)."""
    raw = read_raw(path)
    i = header_index(raw, ("Detalle", "Valor del cargo"))
    if i is None:
        raise ValueError("%s：找不到表头（Detalle / Valor del cargo）" % os.path.basename(path))
    df = frame_below(raw, i)
    out = pd.DataFrame(index=df.index)
    for src, dst in INV_COLS.items():
        if src in df.columns and dst not in out.columns:
            out[dst] = df[src]
    for c in set(INV_COLS.values()):
        if c not in out.columns:
            out[c] = None
    out = out[out["detalle"].map(text) != ""].copy()
    for c in INV_ID_COLS:
        out[c] = out[c].map(idstr)
    for c in INV_NUM_COLS:
        out[c] = out[c].map(num)
    for c in ("detalle", "deducted", "charge_state", "discount_reason", "title"):
        out[c] = out[c].map(text)
    out["charge_date"] = out["charge_date"].map(to_date)
    out["sale_date"] = out["sale_date"].map(to_date)
    name = os.path.basename(path)
    out["file"] = name
    out["doc"] = "NC" if "Notas_Credito" in name else "FAC"
    out["bill_month"] = file_month(name)
    out["row"] = out.index + i + 2           # 1-based row in the sheet
    known = [c for c in ("Cargo que bonifica", "Cargo que anula", "Costo fijo", "Número de kit",
                         "Código ML", "Canal de Venta") if c in df.columns]
    dates = [d for d in out["charge_date"] if d]
    facts = {"name": name, "doc": out["doc"].iloc[0] if len(out) else ("NC" if "Notas" in name else "FAC"),
             "bill_month": file_month(name), "updated": updated_at(raw), "rows": len(out),
             "header_row": i + 1, "first": min(dates) if dates else None,
             "last": max(dates) if dates else None, "columns": known, "open": bill_open(raw, i)}
    return out.reset_index(drop=True), facts


# ------------------------------------------------------------------ invoice payments

def read_pagos(path):
    """Reporte_Pagos_Facturas -> (payments & credit notes, payment detail, facts)."""
    xl = pd.ExcelFile(path)
    summ = det = None
    upd = ""
    for sh in xl.sheet_names:
        raw = pd.read_excel(path, sheet_name=sh, header=None, dtype=object)
        upd = upd or updated_at(raw)
        i = header_index(raw, ("Tipo de pago", "Importe total"))
        if i is not None:
            df = frame_below(raw, i)
            df = df[df["Tipo de pago"].map(text) != ""]
            summ = pd.DataFrame({
                "pay_no": df["Número de pago"].map(idstr) if "Número de pago" in df else "",
                "nc_no": df["Número de nota de crédito"].map(text) if "Número de nota de crédito" in df else "",
                "type": df["Tipo de pago"].map(text),
                "date": df.get("Fecha de pago / Emisión de NC", pd.Series(dtype=object)).map(to_date),
                "state": df.get("Estado", pd.Series(dtype=object)).map(text),
                "total": df["Importe total"].map(num),
                "applied": df.get("Importe aplicado a este mes", df["Importe total"]).map(num),
                "other_month": df.get("Importe aplicado a otro mes", pd.Series(dtype=object)).map(num),
            })
            continue
        i = header_index(raw, ("Número del cargo", "Parte del pago aplicado a cargos"))
        if i is not None:
            df = frame_below(raw, i)
            df = df[df["Número del cargo"].map(idstr) != ""]
            det = pd.DataFrame({
                "pay_no": df["Número del pago"].map(idstr),
                "pay_date": df["Fecha del pago"].map(to_date),
                "applied": df["Parte del pago aplicado a cargos"].map(num),
                "charge_no": df["Número del cargo"].map(idstr),
                "detalle": df["Descripción del cargo"].map(text),
                "charge_date": df["Fecha del cargo"].map(to_date),
            })
    if summ is None:
        summ = pd.DataFrame(columns=["pay_no", "nc_no", "type", "date", "state", "total", "applied",
                                     "other_month"])
    if det is None:
        det = pd.DataFrame(columns=["pay_no", "pay_date", "applied", "charge_no", "detalle", "charge_date"])
    name = os.path.basename(path)
    return summ.reset_index(drop=True), det.reset_index(drop=True), {
        "name": name, "bill_month": file_month(name), "updated": upd}


# ------------------------------------------------------------------ sales export (SKU lookup, status)

VENTAS_COLS = {
    "# de venta": "sale_no", "Fecha de venta": "sale_date_txt", "Estado": "estado",
    "Descripción del estado": "estado_desc", "Paquete de varios productos": "pack_flag",
    "Pertenece a un kit": "kit", "Unidades": "units", "SKU": "sku", "# de publicación": "listing",
    "Título de la publicación": "title", "Variante": "variant",
    "Forma de entrega": "delivery", "Fecha en camino": "shipped_at", "Fecha entregado": "delivered_at",
    "Revisado por Mercado Libre": "reviewed", "Fecha de revisión": "review_date",
    "Dinero a favor": "money_to", "Resultado": "result", "Destino": "destination",
    "Motivo del resultado": "result_reason", "Reclamo abierto": "claim_open",
    "Reclamo cerrado": "claim_closed", "Con mediación": "mediation",
}


# money columns this report takes (headers end in the currency, "(MXN)"; the discount
# column has none).
# A pack's money is on its parent row; its items carry only units and unit price.
VENTAS_MONEY = {"Cargo por venta e impuestos": "v_charge", "Descuentos y bonificaciones": "v_discount"}


def read_ventas(path):
    """The sales export -> (rows, facts). One row per sheet row; `kind` is
    'parent' (multi-product pack head: no SKU, carries the pack id), 'child'
    (an item of the pack above it) or 'single'. Money columns: VENTAS_MONEY."""
    xl = pd.ExcelFile(path)
    sheet = next((s for s in xl.sheet_names if s.lower().startswith("ventas")), xl.sheet_names[0])
    raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=object)
    i = header_index(raw, ("# de venta", "Estado"))
    if i is None:
        raise ValueError("%s：找不到表头（# de venta）" % os.path.basename(path))
    df = frame_below(raw, i)
    out = pd.DataFrame(index=df.index)
    for src, dst in VENTAS_COLS.items():
        out[dst] = df[src] if src in df.columns else None
    unit_col = next((c for c in df.columns if c.startswith("Precio unitario de venta")), None)
    out["unit_price"] = df[unit_col].map(num) if unit_col else 0.0
    for prefix, dst in VENTAS_MONEY.items():
        col = next((c for c in df.columns if c.startswith(prefix)), None)
        out[dst] = df[col].map(num) if col else 0.0
    out = out[out["sale_no"].map(idstr) != ""].copy()
    out["sale_no"] = out["sale_no"].map(idstr)
    out["listing"] = out["listing"].map(idstr)
    for c in ("estado", "estado_desc", "pack_flag", "kit", "sku", "title", "variant", "delivery",
              "shipped_at", "delivered_at", "reviewed", "review_date", "money_to", "result",
              "destination", "result_reason", "claim_open", "claim_closed", "mediation"):
        out[c] = out[c].map(text)
    out["units"] = out["units"].map(num)
    out["sale_date"] = out["sale_date_txt"].map(spanish_date)
    kinds, parents = [], []
    left, parent = 0, ""
    for _, r in out.iterrows():
        m = re.match(r"Paquete de (\d+) productos", r["estado"])
        if m and not r["sku"]:
            kinds.append("parent")
            parents.append("")
            left, parent = int(m.group(1)), r["sale_no"]
        elif left > 0:
            kinds.append("child")
            parents.append(parent)
            left -= 1
        else:
            kinds.append("single")
            parents.append("")
    out["kind"] = kinds
    out["parent"] = parents
    dates = [d for d in out["sale_date"] if d]
    return out.reset_index(drop=True), {
        "name": os.path.basename(path), "sheet": sheet, "header_row": i + 1, "rows": len(out),
        "first": min(dates) if dates else None, "last": max(dates) if dates else None}


# ------------------------------------------------------------------ MercadoPago (evidence)

def _read_csv(path):
    for enc in ("utf-8-sig", "latin-1"):
        try:
            return pd.read_csv(path, sep=None, engine="python", dtype=str, encoding=enc)
        except (UnicodeDecodeError, ValueError):
            continue
    return pd.DataFrame()


def _col(df, *prefixes):
    for p in prefixes:
        for c in df.columns:
            if c.startswith(p):
                return c
    return None


def read_collection(path):
    """Cobros: per payment — order id, seller SKU, listing, status, refunds, shipment."""
    df = _read_csv(path)
    g = lambda *p: df[_col(df, *p)] if _col(df, *p) else pd.Series([None] * len(df), index=df.index)
    out = pd.DataFrame({
        "op_id": g("Número de operación de Mercado Pago").map(idstr),
        "order_id": g("Número de venta en Mercado Libre").map(idstr),
        "sku": g("SKU Producto").map(text),
        "item_id": g("Identificador de producto").map(idstr),
        "status": g("Estado de la operación").map(text),
        "status_detail": g("Detalle del estado de la operación").map(text),
        "shipment": g("Estado del envío").map(text),
        "refunded": g("Monto devuelto").map(num),
        "amount": g("Valor del producto").map(num),
        "date": g("Fecha de compra").map(text),
    })
    return out[out["op_id"] != ""].reset_index(drop=True), {"name": os.path.basename(path), "rows": len(out)}


def read_settlement(path):
    """settlement_v2: money movements per payment (SOURCE_ID = the invoice's Pago)."""
    df = _read_csv(path)
    out = pd.DataFrame({
        "source_id": df.get("SOURCE_ID", pd.Series(dtype=str)).map(idstr),
        "type": df.get("TRANSACTION_TYPE", pd.Series(dtype=str)).map(text),
        "amount": df.get("TRANSACTION_AMOUNT", pd.Series(dtype=str)).map(num),
        "taxes": df.get("TAXES_AMOUNT", pd.Series(dtype=str)).map(num),
        "date": df.get("TRANSACTION_DATE", pd.Series(dtype=str)).map(lambda s: text(s)[:10]),
    })
    return out[out["source_id"] != ""].reset_index(drop=True), {"name": os.path.basename(path),
                                                                "rows": len(out)}


# ------------------------------------------------------------------ Full / stock (SKU attribution)

def read_cargos_full(path):
    """Reporte_Cargos_Full: every sheet; charge number -> SKU / listing where the sheet has them."""
    xl = pd.ExcelFile(path)
    frames = []
    for sh in xl.sheet_names:
        raw = pd.read_excel(path, sheet_name=sh, header=None, dtype=object)
        i = header_index(raw, ("Detalle", "Monto del cargo"))
        if i is None:
            continue
        df = frame_below(raw, i)
        cno = _col(df, "N° de cargo", "Nº de cargo", "N° del cargo", "Nº del cargo")
        if cno is None:
            continue
        lst = _col(df, "N° de publicación", "Nº de publicación", "Número de publicación")
        out = pd.DataFrame({
            "charge_no": df[cno].map(idstr),
            "detalle": df["Detalle"].map(text),
            "amount": df["Monto del cargo"].map(num),
            "sku": df["SKU"].map(text) if "SKU" in df.columns else "",
            "inv_code": df["Código ML"].map(text) if "Código ML" in df.columns else "",
            "listing": df[lst].map(listing_digits) if lst else "",
            "sheet": sh,
        })
        frames.append(out[out["charge_no"] != ""])
    if not frames:
        return pd.DataFrame(columns=["charge_no", "detalle", "amount", "sku", "inv_code", "listing",
                                     "sheet"]), {"name": os.path.basename(path)}
    return pd.concat(frames, ignore_index=True), {"name": os.path.basename(path)}


def read_storage(path):
    """Costos_por_servicio_almacenamiento → Detalle: per SKU and day, the storage cost."""
    raw = pd.read_excel(path, sheet_name="Detalle", header=None, dtype=object)
    i = header_index(raw, ("SKU", "Código ML"))
    if i is None:
        return pd.DataFrame(columns=["sku", "inv_code", "listing", "title", "day", "cost"]), {
            "name": os.path.basename(path)}
    df = frame_below(raw, i)
    day_cols = [c for c in df.columns if re.match(r"^\d{2}/\d{2}/\d{4}$", c)]
    df = df[df["Código ML"].map(text).str.match(r"^[A-Z]{3,5}\d{3,}$")]
    lst = _col(df, "Número de publicación")
    rows = []
    for _, r in df.iterrows():
        for c in day_cols:
            v = num(r[c])
            if v:
                rows.append((text(r["SKU"]), text(r["Código ML"]), listing_digits(r[lst]) if lst else "",
                             text(r.get("Producto")), datetime.datetime.strptime(c, "%d/%m/%Y").date(), v))
    out = pd.DataFrame(rows, columns=["sku", "inv_code", "listing", "title", "day", "cost"])
    days = [datetime.datetime.strptime(c, "%d/%m/%Y").date() for c in day_cols]
    return out, {"name": os.path.basename(path), "first": min(days) if days else None,
                 "last": max(days) if days else None}


STOCK_SHEETS = ("Buena calidad", "Para impulsar ventas", "Para poner en venta", "Para evitar descarte")


def read_stock_general(path):
    """stock_general_full: a snapshot; per SKU and state sheet, the units (last number on the row)."""
    xl = pd.ExcelFile(path)
    rows, snap = [], ""
    for sh in xl.sheet_names:
        raw = pd.read_excel(path, sheet_name=sh, header=None, dtype=object)
        for k in range(min(8, len(raw))):
            for v in raw.iloc[k].tolist():
                t = text(v)
                if t.startswith("Actualizado"):
                    snap = snap or t
        if sh not in STOCK_SHEETS:
            continue
        i = header_index(raw, ("Código ML", "SKU"))
        if i is None:
            continue
        hdr = [text(v) for v in raw.iloc[i].tolist()]
        ci, si = hdr.index("Código ML"), hdr.index("SKU")
        pi = next((k for k, h in enumerate(hdr) if h.startswith("Publicaciones")), None)
        ti = hdr.index("Producto") if "Producto" in hdr else None
        s30 = next((k for k, h in enumerate(hdr) if h.startswith("Ventas últimos 30 días (u") or
                    h.startswith("Unidades vendidas")), None)
        for k in range(i + 1, len(raw)):
            vals = raw.iloc[k].tolist()
            code = text(vals[ci])
            if not re.match(r"^[A-Z]{3,5}\d{3,}$", code):
                continue
            nums = [v for v in vals if isinstance(v, (int, float)) and v == v]
            units = float(nums[-1]) if nums else 0.0
            rows.append((text(vals[si]), code, text(vals[pi]) if pi is not None else "",
                         text(vals[ti]) if ti is not None else "", sh, units,
                         num(vals[s30]) if s30 is not None else 0.0))
    out = pd.DataFrame(rows, columns=["sku", "inv_code", "listings", "title", "state", "units", "sold30"])
    return out, {"name": os.path.basename(path), "snapshot": snap}


def read_conciliation(path):
    """conciliation_*: stock movements over a date range, per SKU."""
    raw = read_raw(path)
    i = header_index(raw, ("SKU", "Código ML", "Ingresadas"))
    if i is None:
        return pd.DataFrame(), {"name": os.path.basename(path)}
    hdr = [text(v) for v in raw.iloc[i].tolist()]
    stock_cols = [k for k, h in enumerate(hdr) if h.startswith("Stock total almacenado")]
    want = {"Ingresadas": "in", "Vendidas": "sold", "Ventas canceladas": "cancelled",
            "Devueltas por comprador": "returned", "Retiradas": "withdrawn"}
    rows = []
    for k in range(i + 1, len(raw)):
        vals = raw.iloc[k].tolist()
        code = text(vals[hdr.index("Código ML")])
        if not re.match(r"^[A-Z]{3,5}\d{3,}$", code):
            continue
        rec = {"sku": text(vals[hdr.index("SKU")]), "inv_code": code,
               "listings": text(vals[next((j for j, h in enumerate(hdr) if h.startswith("Publicaciones")),
                                          hdr.index("SKU"))]),
               "title": text(vals[hdr.index("Producto")]) if "Producto" in hdr else "",
               "start": num(vals[stock_cols[0]]) if stock_cols else 0.0,
               "end": num(vals[stock_cols[-1]]) if stock_cols else 0.0}
        for h, key in want.items():
            rec[key] = num(vals[hdr.index(h)]) if h in hdr else 0.0
        rows.append(rec)
    labels = [hdr[k] for k in stock_cols]
    return pd.DataFrame(rows), {"name": os.path.basename(path),
                                "range": "%s → %s" % (labels[0][-5:], labels[-1][-5:]) if labels else ""}


def read_returns(path):
    """Returns_<date>.xlsx (Triages), Spanish or English headers."""
    raw = read_raw(path)
    i = header_index(raw, ("Order", "SKU")) if header_index(raw, ("Order", "SKU")) is not None else \
        header_index(raw, ("Número de orden", "SKU"))
    if i is None:
        return pd.DataFrame(columns=["order", "inv_code", "review_date", "product_status", "review_result",
                                     "money_status"]), {"name": os.path.basename(path), "rows": 0}
    df = frame_below(raw, i)
    pick = lambda *names: next((df[n] for n in names if n in df.columns), pd.Series([None] * len(df)))
    out = pd.DataFrame({
        "order": pick("Order", "Número de orden").map(idstr),
        "inv_code": pick("SKU").map(text),
        "review_date": pick("Review date", "Fecha de revisión").map(text),
        "product_status": pick("Product status", "Estado del producto").map(text),
        "review_result": pick("Review result", "Resultado de la revisión").map(text),
        "money_status": pick("Money status", "Estado del dinero").map(text),
    })
    out = out[out["order"] != ""].reset_index(drop=True)
    return out, {"name": os.path.basename(path), "rows": len(out)}
