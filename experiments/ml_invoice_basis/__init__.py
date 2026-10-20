# -*- coding: utf-8 -*-
"""Experiment: a MercadoLibre P&L on the invoice basis (Facturación + Notas de
crédito of the month) next to the current order basis (Ventas), and how stable
each basis is across several downloads of the same month.

    python -m experiments.ml_invoice_basis --store <name> --month YYYY-MM [--folder DIR]
    python -m experiments.ml_invoice_basis.stability --store <name> --month YYYY-MM [--folders DIR ...]

See README.md in this folder. Not part of the monthly run.
"""
