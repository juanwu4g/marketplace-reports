# Downloaded Forms — Catalogue (Liverpool Mexico)

Every form this platform needs from Liverpool (El Puerto de Liverpool) to build
**last month's** financial report: where it is, how to get it, and what is inside.
Same structure as `platforms/falabella/FORMS.md` and `platforms/amazon/FORMS.md`.

**This document grows.** Keep the "verified" date honest.

- Seller: SHEN MING SA DE CV, supplier number 0000163073, shop "Vita Hogar".
- Ziniao stores (platform name `自定义平台`, tag `MX 商超`, proxy 216.238.92.32):
  - `Liverpool Order 子账号(Nikki)`: **the store**. Its marketplace session stays
    signed in inside Ziniao. It opens the supplier portal in the same browser and
    signs in as P00163073 with config.json's `portal_login`.
  - `Liverpool Report`: the supplier-portal user. **Not used** (owner, 2026-10-27).
  - `Liverpool Order Admin`, `Liverpool Order 子账号(Simon)`: other marketplace
    users of the same seller, not used.
- Last verified: **2026-10-26**, by hand-driven probes in the test VM (September
  2026, every form).

---

## Index

| # | Form id | Site | Page | Period | File |
|---|---|---|---|---|---|
| 1 | ledger | supplier portal | Facturación y Pagos → Pagos MKP → **3.- Cobros y Pagos Marketplace** | posting date in the month | `LV_mayor_<YYYYMM>.xlsx` |
| 2 | sales | supplier portal | Pagos MKP → **9.- Ventas Marketplace** | date in the month | `LV_ventas_<YYYYMM>.xlsx` |
| 3 | payments | supplier portal | Pagos MKP → **2.- Detalle de Pagos** | payments of the month + open items | `LV_pagos_<YYYYMM>.xlsx`, `LV_pagos_pendientes_<YYYYMM>.xlsx` |
| 4 | invoices | supplier portal | Pagos MKP → **7.- Notas de Cargo** | folders `<YYYY>/<MM>` of the month and the next one | `Notas_Cargo_<YYYYMM>/*.xml` + `*.pdf` |
| 5 | orders | marketplace | Pedidos → Pedidos y facturas → Descargar → XLSX | the last four months, as the page offers | `LV_pedidos_<YYYYMMDD>.xlsx` |
| 6 | logistics | marketplace | Pedidos → Costos logísticos → Filtros → Fecha de cobro | charge date in the month | `LV_costos_logisticos_<YYYYMM>.xlsx` |
| 7 | balance | marketplace | Mi tienda → Balance de cuenta → Filtros → Fecha de registro | registration date in the month | `LV_balance_<YYYYMM>.xlsx` |

Every download is immediate (no "we'll prepare it"): nothing is ever owed.

**Timing for last month:** the ledger and the sales report are complete for M from
the first day of M+1 (Liverpool posts a sale on its order day and the fees on their
charge day). The fee CFDIs of M's last week are issued in M+1. Running on or after
the 5th of M+1 also catches the first payment of M+1 in "Detalle de Pagos".

## Signing in — read before touching the routes

- **Supplier portal** `https://proveedores.liverpool.com.mx/` (SAP NetWeaver).
  - **Akamai first:** open `https://entradaunica.liverpool.com.mx/portal` first;
    without its cookies (`ak_bmsc`, `_abck`, `bm_sz`) the portal answers "Access
    Denied". Seen 2026-10-23 and 2026-10-26.
  - **The sign-in form:** fields `#logonuidfield` / `#logonpassfield` (names
    `j_username` / `j_password`), submit button name `uidPasswordLogon`.
  - **Who types the password:** Ziniao never fills saved passwords in an
    automated (WebDriver) browser, so the routes type config.json's `portal_login`
    into this form (owner's decision 2026-10-27). Without it the store is "needs
    login" (`NeedsLogin` → PAGE_BLOCKED).
  - **Sign-in cookies:** `MYSAPSSO2` and `JSESSIONID` are session cookies; every
    run signs in again.
  - **Signed-in marker:** "Salir del sistema" plus "Bienvenido SHENMING SA DE CV".
  - **Opening the portal directly:** the deep link `/irj/portal` gave "Access
    Denied" once; open the root.
- **Payment pages:** the pages under Pagos MKP (except Notas de Cargo) are Web
  Dynpro apps on `https://itsproves.liverpool.com.mx/sap/bc/webdynpro/sap/…`,
  loaded into the portal's frame `isolatedWorkArea` (inside the frame `Área
  interna de desktop`).
  - **Ziniao blocked that host** until the owner's whitelist request was
    approved (2026-10-26). The approval takes effect only after the **Ziniao
    client** restarts, not the store.
  - **Symptom when blocked:** "Chromium ha bloqueado esta página" in the frame, or
    Ziniao's `stop.html` page opened at the top level.
- **Marketplace** `https://marketplace.liverpool.com.mx` (React, MUI).
  - **The session:** it lives in the Ziniao store. A signed-out store lands on
    `login-entradaunica.liverpool.com.mx/u/login` (Auth0, email + password),
    which is a person's job.
  - **Ziniao's extra handle:** Ziniao opens with an extra handle (an omnibox
    pop-up), and using it makes the site say "Este sitio no está diseñado para
    … móvil". Switch to the first window whose size can be read (`real_window`).
- **Entrada Única** (the "new website" of the 2026-10-22 announcement "¡Nos
  actualizamos para ti!") is the single sign-on portal. With the marketplace user
  it offers only Marketplace and Catalogación; the payment pages are still the SAP
  portal's.

## 1. ledger — "Saldos Proveedor" (Cobros y Pagos Marketplace)

App `ZFI_GL_CF_SALDOSPROV`. Steps:
1. Choose **Todas las Partidas**.
2. Enter **Fecha de Contabilización** from `01.09.2026` to `30.09.2026` (the two
   text inputs after that label; Web Dynpro ids change, so locate them by the
   label).
3. Click the green tick at the top left (a `div[ct=B]` whose `lsdata` contains
   "Concluido").
4. Click "Export" → "Exportar a Microsoft Excel".

Columns: ID, Soc., **Lib.mayor**, Acreedor, **Nº doc.**, Año, **Asignación**, BP,
CT, **Clase doc.**, VP, **Texto**, **II** (tax code), **Doc.comp.**, **Referencia**,
Fe.pago, **Fe.contab.**, Fecha doc., **Importe Debe/Haber**, ML, Impte.mon.extranj.,
Mon., Compens. The last row is the total (no Nº doc.).

Two accounts (`catalog.py` registers every kind of line):

| Lib.mayor | Clase | Texto / II | Meaning |
|---|---|---|---|
| 20040002 (Liverpool owes the seller) | YV | `F/ …` D2 | sale price paid by the buyer (credit) |
| | YV | `F/` C2, own document, same Referencia as the sale | Liverpool's promotion paid back to the seller |
| | YV | `COM …` Q6 / Q8 | IVA withheld 8 % / ISR withheld 2.5 % of the price net of IVA |
| | YD | `DEV …` | cancellation or return of a posted sale |
| | YD | `F/ …` Q6 / Q8 / C5 | withholding / commission given back on it |
| | RW | `Ajuste log …` | logistics adjustment of an earlier month |
| | ZP | — | payments (pairs that net to 0 = a payment sent back by the bank) |
| 11900003 (the seller owes Liverpool) | YV | `<store>/Comisión % DILISA` | commission (~17 % of the price), same Nº doc. as the sale |
| | YV | `<store>/Sobrecomis % T.Exter` | surcharge when the buyer pays with a non-Liverpool card |
| | DY | `Rec Logística dd/mm/yyyy` | shipping charges of a day; the Nº doc. is the "Doc contable" of the logistic-cost list |
| | DY | `Balance logistico <mes>` | extra shipping (Balance de cuenta) |
| | DY | `AJUSTE MSI …`, `Ajuste log …`, `Suscripcion MKP …` | interest-free instalment cost, logistics adjustments, the monthly fee (600) |
| | ZP | `PAGO`, `RECHAZO BANCO`, … | the fees settled with the payments |

September 2026: 707 lines. Every amount of the report is −1 × Importe.

## 2. sales — "Reporte Market Place" (Ventas Marketplace)

App `zrt_wd_repliq`. Steps:
1. Enter **Fecha** from / to (the first two text inputs after that label).
2. Click **Ejecutar**.
3. Click "Export" → Excel.

One row per order of a sale or return document. Columns:
- Almacén
- Fecha (= the ledger's posting date)
- Proveedor
- **Artículo** (Liverpool SKU number = the order list's "No de SKU Liverpool")
- **Nº documento** (= the ledger's Nº doc.)
- **Referencia** (its last ten digits are the marketplace order number)
- Venta con impuestos
- Descuento (Liverpool's promotion)
- Venta con descuento
- Comisión Dilisa
- Pago neto

September 2026: 140 rows.
- **Orders:** 139 matched the order list. The other one was the return of an order
  older than the list's four months.
- **Commission:** the total equals the ledger's net commission to the cent.

## 3. payments — Detalle de Pagos

App `zfi_ap_inf_pago_prov`. Two tables, each with a "Generar Archivo en Excel"
menu → "Exportar a Microsoft Excel":
- **Pagos Pendientes** (open items).
- **Pagos realizados**: enter "Fecha de pago (año)" and "(mes)", then click **Buscar**.

A payment's lines (button "Detalles", after clicking its Doc.comp.) are the payable
account's side only. September 2026: one payment of 423,244.53 on 28.09, with
820 detail lines.

## 4. invoices — Notas de Cargo

A Knowledge Management folder on the portal itself (not blocked):
`/irj/go/km/docs/factura_elec/Nota_Cargo/<user>/<YYYY>/<MM>/`, 15 files per
page. Each file is fetched in the page with `fetch(…, {credentials:'include'})`.
- **What they are:** Liverpool's CFDIs to the seller.
- **Series:** MCOM (weekly commissions; E = credit notes for returns), MCUO
  (subscription), MLOG (logistics).
- **September 2026:** 40 files (20 XML).

## 5. orders — Pedidos

`/orders` → **Descargar** → XLSX. The page shows the last four months ("Visualizando
últimos cuatro meses"); the export follows. Sheet `OrderReport`, 85 columns.
- **Kept:** dates, order id, status, Liverpool SKU, seller SKU ("SKU de la
  oferta"), product name, quantities and prices, Liverpool promotion, refund flags,
  delivery and pick-up dates.
- **Dropped on load:** the buyer's name, address and phone.

## 6. logistics — Costos logísticos

`/orders/admin/logistic-costs`. Steps:
1. **Filtros**.
2. Enter `charged_at_start` / `charged_at_end`, typing DD/MM/YYYY digit by digit
   (MUI date fields).
3. **Aplicar filtros**.
4. **Descargar tabla** → XLSX.

Sheet `LogisticCostReportSeller`: Nombre paquetera, CR Liverpool, Pedido, Tipo,
Fecha Cobro, Costos Logisticos, Costo de seguro, Fecha envío SAP, Estado de cobro,
**Doc contable**. September 2026: 161 rows, 32,782.00. Every Doc contable matches a
"Rec Logística" document of the ledger to the cent; the ledger has 12 more
shipping documents (5,511.50) that are not in this list.

## 7. balance — Balance de cuenta

`/stores/balance` (Mi tienda → Balance de cuenta). Steps:
1. **Filtros**.
2. Enter Fecha de registro (`input[name=from]` / `[name=to]`, the first pair).
3. **Aplicar filtros**.
4. **Descargar** → XLSX.

If the Motivo dropdown is opened, close it first: clicking elsewhere while it is
open does nothing. Sheet `AdjustmentLogisticCostReportSel`: Id ajuste, Fecha de
registro, Estatus, **Remisión** (order id), Monto, Motivo (COMPLEMENTARY, DIGITAL,
compensation), Comentario, Clase de doc, Doc contable, Fecha cobro real, Pedido
SAP, Contra recibo.
