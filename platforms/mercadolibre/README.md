# MercadoLibre — platform notes

**Audience:** developers changing the MercadoLibre downloads or reports. The
form-by-form detail is in `FORMS.md` (same folder); operators read the root
`README.md`.

## Countries (sites)

Per-country facts live in `meli_site.py`; a store's country comes from
`config.json` → `stores.regions` (default MX).

| | Mexico `mx` | Chile `cl` | Peru `pe` |
|---|---|---|---|
| Ziniao platform name | MercadoLibre-墨西哥-本土 | MercadoLibre-智利-本土 | MercadoLibre前台-秘鲁 |
| seller centre | vendedores.mercadolibre.com.mx | vendedores.mercadolibre.cl | vendedores.mercadolibre.com.pe |
| MercadoPago | mercadopago.com.mx | mercadopago.cl | mercadopago.com.pe |
| UI language required | es_MX | es_CL | es_PE |
| currency | MXN, 2 decimals | CLP, **no decimals** | PEN, 2 decimals |
| timezone | America/Mexico_City | America/Santiago | America/Lima |
| sales sheet | "Ventas MX", columns "(MXN)" | "Ventas CL", "(CLP)" | "Ventas PE", "(PEN)" |
| monthly bill | calendar month | **per-store cut day** (e.g. 8th→7th, 24th→23rd); charges are assigned by charge date | calendar month |
| Full warehouse (stock reports) | yes | yes | **no** — stock pages are 404; not downloaded, not expected |
| tax | IVA 16%; platform withholds IVA/ISR | IVA 19%; nothing withheld | IGV 18%; nothing withheld; fee invoices are boletas (`B001-…`) |

## Quirks found in production (keep them in mind)

- **Peru spells September "Setiembre"** in the seller centre (bill cards, detail
  page). Month regexes use `Sep?tiembre`. MercadoPago's date picker still says
  "septiembre", so MercadoPago month names are not switched per site.
- **Peru's sales-export notification clears** after a while: the export is
  collected right after requesting it (`ventas_collect_now` in `meli_site.py`).
- **Chile's returns menu** says "Reporte de devoluciones al centro de
  distribución" (Mexico: "… a la bodega"); both are accepted.
- **New accounts**: MercadoPago greys out dates before the account existed; the
  start date moves to the first selectable day.
- **Buyer-paid shipping** (Chile, Peru) and the Flex delivery bonus (Peru) are
  separate MercadoPago payments, so the per-order MercadoPago comparison shows
  explained differences.
- **Platform says the sales export is incomplete** ("No pudimos mostrar algunos
  datos"): the report is still built, with a warning.
- The machine clock is China time; "today" for the platform is computed in the
  site's timezone (`core/site_time.py`).

## Forms not downloaded on purpose

- MercadoPago account statement ("Estados de saldos y movimientos"): the
  "Generar" button is often not found (always, so far, in Chile and Peru); the
  report does not use it.
- `BILL_MP_V3` (fifth billing file): asks for a format in a dialog the script
  does not answer; the report does not use it.
- Both are recorded with a cause but do not count as "report missing data"
  (`NOT_IN_REPORT` in `core/download_diag.py`).

## Accounting rules

The rules, the ten sheets and the per-country differences are documented for
operators and finance in the root `README.md` §3.6 (countries) and §4 (sheets).
Open accountant questions per country are listed in the run reports' ⑩ sheet.

## Tests

`tests/test_cl_*.py`, `tests/test_pe_*.py` (synthetic data) and
`tests/test_golden.py` (golden: real Mexican, Chilean and Peruvian store-months
rebuilt cell by cell against the accepted workbooks; data in `~/report-golden/`,
outside git).
