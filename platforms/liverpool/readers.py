# -*- coding: utf-8 -*-
"""Readers of Liverpool's exports (FORMS.md). Each returns a DataFrame with short
column names, amounts as floats, dates as Timestamps, totals rows dropped.

The order list carries customer data (name, address, phone): only the columns in
ORDER_COLS are kept, the rest is dropped on load and never reaches a workbook.
"""
import glob
import os
import re
import warnings
import xml.etree.ElementTree as ET

import pandas as pd

# Liverpool's exports carry no default cell style; openpyxl says so for every file.
warnings.filterwarnings("ignore", message="Workbook contains no default style", category=UserWarning)

# names the downloads are saved under (forms.py), per month stamp YYYYMM
LEDGER = "LV_mayor_%s.xlsx"
SALES = "LV_ventas_%s.xlsx"
PAYMENTS = "LV_pagos_%s.xlsx"
PENDING = "LV_pagos_pendientes_%s.xlsx"
INVOICE_DIR = "Notas_Cargo_%s"
ORDERS = "LV_pedidos_*.xlsx"
LOGISTICS = "LV_costos_logisticos_%s.xlsx"
BALANCE = "LV_balance_%s.xlsx"


def num(s):
    """'1,099.00' / '1,099.00-' / 1099 / NaN -> float (SAP writes a trailing minus)."""
    def one(x):
        if x is None or (isinstance(x, float) and x != x):
            return 0.0
        if isinstance(x, (int, float)):
            return float(x)
        t = str(x).replace(",", "").replace("$", "").replace("MXN", "").strip()
        if not t or t.lower() == "nan":
            return 0.0
        neg = t.endswith("-")
        try:
            v = float(t.rstrip("-"))
        except ValueError:
            return 0.0
        return -v if neg else v
    return s.map(one)


def day(s, fmt=None):
    """Dates as Timestamps (time of day dropped). Liverpool writes '30/06/2026 - 03:42 PM'
    on the marketplace and Excel dates or '29.09.2026' on the portal."""
    def one(x):
        if x is None or (isinstance(x, float) and x != x):
            return pd.NaT
        if isinstance(x, pd.Timestamp):
            return x.normalize()
        t = str(x).strip()
        for f, n in (("%d/%m/%Y", 10), ("%d.%m.%Y", 10), ("%Y-%m-%d", 10), ("%d/%m/%y", 8)):
            try:
                return pd.Timestamp(pd.to_datetime(t[:n], format=f))
            except (ValueError, TypeError):
                continue
        return pd.NaT
    return s.map(one)


def text(s):
    return s.fillna("").astype(str).str.replace(r"\s+", " ", regex=True).str.strip()


def col(x, name, default=""):
    """Column `name`, or a column of `default` when the export lacks it."""
    return x[name] if name in x.columns else pd.Series(default, index=x.index)


def ids(s):
    """Document / order numbers as text, without '.0' and surrounding blanks."""
    return text(s).str.replace(r"\.0$", "", regex=True)


# ------------------------------------------------------------------ portal

def read_ledger(path):
    """'Saldos Proveedor' (Cobros y Pagos Marketplace, Todas las Partidas)."""
    x = pd.read_excel(path, dtype=object)
    x = x[x["Nº doc."].notna()].copy()
    return pd.DataFrame({
        "account": ids(x["Lib.mayor"]), "doc": ids(x["Nº doc."]), "cls": text(x["Clase doc."]),
        "text": text(x["Texto"]), "tax": text(x["II"]), "ref": ids(x["Referencia"]),
        "assign": text(x["Asignación"]), "clearing": ids(x["Doc.comp."]),
        "post": day(x["Fe.contab."]), "doc_date": day(x["Fecha doc."]),
        "cleared": day(x["Compens."]), "v": num(x["Importe Debe/Haber"]),
    }).reset_index(drop=True)


def read_sales(path):
    """'Reporte Market Place' (Ventas Marketplace): one row per order of a sale or
    return document; the order number is the last ten digits of Referencia."""
    x = pd.read_excel(path, dtype=object)
    x = x[x["Nº documento"].notna()].copy()
    ref = ids(x["Referencia"])
    return pd.DataFrame({
        "doc": ids(x["Nº documento"]), "date": day(x["Fecha"]), "store_lv": ids(x["Almacén"]),
        "article": ids(x["Artículo"]).str.lstrip("0"), "ref": ref, "order": ref.str[-10:],
        "gross": num(x["Venta con impuestos"]), "discount": num(x["Descuento"]),
        "net": num(x["Venta con descuento"]), "commission": num(x["Comisión Dilisa"]),
        "net_pay": num(x["Pago neto"]),
    }).reset_index(drop=True)


def read_payments(path):
    """'Pagos realizados' of Detalle de Pagos: one row per payment."""
    x = pd.read_excel(path, dtype=object)
    x = x[x["Doc.comp."].notna()].copy()
    return pd.DataFrame({"clearing": ids(x["Doc.comp."]), "paid": day(x["Fecha pago"]),
                         "amount": num(x["Importe"]),
                         "kind": text(col(x, "Tipo de pago", ""))}
                        ).reset_index(drop=True)


def read_invoices(folder):
    """The fee CFDIs (XML) of one Notas de Cargo folder: one row per document."""
    rows = []
    for p in sorted(glob.glob(os.path.join(folder, "*.xml"))):
        try:
            root = ET.parse(p).getroot()
        except ET.ParseError:
            continue
        a = root.attrib
        concept = next((e for e in root.iter() if e.tag.endswith("Concepto")), None)
        tax = next((e for e in root.iter() if e.tag.endswith("Impuestos")
                    and "TotalImpuestosTrasladados" in e.attrib), None)
        rows.append({"file": os.path.basename(p), "series": a.get("Serie", ""), "folio": a.get("Folio", ""),
                     "date": pd.Timestamp(a.get("Fecha", "")[:10]) if a.get("Fecha") else pd.NaT,
                     "type": a.get("TipoDeComprobante", ""),
                     "subtotal": float(a.get("SubTotal") or 0), "total": float(a.get("Total") or 0),
                     "iva": float(tax.attrib["TotalImpuestosTrasladados"]) if tax is not None else 0.0,
                     "text": concept.attrib.get("Descripcion", "") if concept is not None else ""})
    return pd.DataFrame(rows, columns=["file", "series", "folio", "date", "type", "subtotal", "total",
                                       "iva", "text"])


# ------------------------------------------------------------------ marketplace

# the order list's columns that are kept (everything else — the buyer's name,
# address, phone, e-mail, RFC — is dropped on load)
ORDER_COLS = ["Fecha de creación", "Id del pedido", "Estado", "Canal de venta", "Método de envío",
              "Tipo de artículo", "No de SKU Liverpool", "Nombre del producto", "SKU de la oferta",
              "Información adicional sku", "Marca", "Cantidad", "Precio por unidad", "Subtotal",
              "Promoción Liverpool", "Pedidos con reembolso", "Reembolso parcial", "Reembolso total",
              "Piezas devueltas por indicador", "Fecha de entrega", "Fecha de recepción", "Fulfillment (Si/No)",
              "Total de la orden a pagar al seller"]


def read_orders(path):
    x = pd.read_excel(path, dtype=object)
    x = x[[c for c in ORDER_COLS if c in x.columns]].copy()
    x = x[x["Id del pedido"].notna()]
    out = pd.DataFrame({
        "order": ids(x["Id del pedido"]), "created": day(x["Fecha de creación"]),
        "status": text(x["Estado"]), "channel": text(col(x, "Canal de venta")),
        "article": ids(x["No de SKU Liverpool"]).str.lstrip("0"), "sku": text(x["SKU de la oferta"]),
        "title": text(x["Nombre del producto"]), "variant": text(col(x, "Información adicional sku")),
        "qty": num(x["Cantidad"]), "subtotal": num(x["Subtotal"]),
        "promo": num(col(x, "Promoción Liverpool", 0)),
        "refund_flag": text(col(x, "Pedidos con reembolso", "")),
        "returned_units": num(col(x, "Piezas devueltas por indicador", 0)),
        # home delivery: Fecha de entrega; store pick-up (Click & Collect): Fecha de recepción
        "delivered": day(col(x, "Fecha de entrega", None)).fillna(day(col(x, "Fecha de recepción", None))),
    })
    return out.drop_duplicates(subset=["order", "article"]).reset_index(drop=True)


def read_logistics(path):
    """'Costos logísticos' (DETALLE_COSTOS_LOG): one row per charged order."""
    x = pd.read_excel(path, dtype=object)
    x = x[x["Pedido"].notna()].copy()
    return pd.DataFrame({
        "carrier": text(x["Nombre paquetera"]), "order": ids(x["Pedido"]), "kind": text(x["Tipo"]),
        "charged": day(x["Fecha Cobro"]), "amount": num(x["Costos Logisticos"]),
        "insurance": num(x["Costo de seguro"]), "status": text(x["Estado de cobro"]),
        "doc": ids(x["Doc contable"]),
    }).reset_index(drop=True)


def read_balance(path):
    """'Balance de cuenta' (AdjustmentLogisticCostReportSeller): extra charges."""
    x = pd.read_excel(path, dtype=object)
    x = x[x["Id ajuste"].notna()].copy()
    return pd.DataFrame({
        "id": ids(x["Id ajuste"]), "registered": day(x["Fecha de registro"]),
        "status": text(x["Estatus"]), "order": ids(x["Remisión"]), "amount": num(x["Monto"]),
        "reason": text(x["Motivo"]), "doc": ids(x["Doc contable"]),
    }).reset_index(drop=True)


def month_files(folder, pattern):
    """Matches of `pattern` in the folder, newest name last."""
    return sorted(glob.glob(os.path.join(folder, pattern)))


def order_file_date(path):
    m = re.search(r"(\d{8})", os.path.basename(path))
    return m.group(1) if m else ""
