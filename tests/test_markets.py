# -*- coding: utf-8 -*-
"""Markets = platform x country. Old configs keep working; new ones are explicit;
a store of one platform can never be picked up by another platform's batch."""
import json

import pytest

from core import markets
from core import store_config
from platforms.mercadolibre import meli_site
from platforms.mercadolibre import run_batch

LEGACY = {"default": "BOCINA_SM",
          "allowed": ["BOCINA_SM", "EWTTO_SM", "Tecnohub_CL", "IDG_PE"],
          "batch": ["BOCINA_SM", "EWTTO_SM"],
          "regions": {"Tecnohub_CL": "CL", "IDG_PE": "pe"},
          "batches": {"CL": ["Tecnohub_CL"]}}

NEW = {"default": "BOCINA_SM",
       "allowed": ["BOCINA_SM", "Tecnohub_CL", "DAB_CL-Fala (Coco)", "AIS_CL-Fala"],
       "catalog": {"BOCINA_SM": {"platform": "ML", "country": "MX"},
                   "Tecnohub_CL": {"platform": "ml", "country": "cl"},
                   "DAB_CL-Fala (Coco)": {"platform": "FA", "country": "CL"},
                   "AIS_CL-Fala": {"platform": "FA", "country": "CL"}},
       "batches": {"FA-CL": ["DAB_CL-Fala (Coco)"]}}


def use(monkeypatch, cfg):
    d = json.loads(json.dumps(cfg))
    monkeypatch.setattr(store_config, "_stores", lambda: d)
    return d


# ------------------------------------------------------------------ registry

def test_market_ids_and_labels():
    assert markets.normalize("fa-cl") == "FA-CL"
    assert markets.normalize("CL") == "ML-CL"            # bare country = MercadoLibre
    assert markets.label("FA-CL") == "Falabella 智利"
    assert markets.label("ML-MX") == "美客多 墨西哥"
    assert "FA-CL" in markets.all_markets() and "ML-PE" in markets.all_markets()


def test_unknown_or_unsupported_markets_raise():
    with pytest.raises(markets.UnknownMarket):
        markets.normalize("XX-CL")
    with pytest.raises(markets.UnknownMarket):
        markets.normalize("FA-MX")                       # Falabella has no Mexico (yet)
    with pytest.raises(markets.UnknownMarket):
        markets.normalize("ML-ZZ")


def test_ziniao_platform_names():
    assert markets.ziniao_matches("ML", "MercadoLibre-墨西哥-本土")
    assert markets.ziniao_matches("ML", "MercadoLibre前台-秘鲁")
    assert markets.ziniao_matches("FA", "falabella-智利")
    assert not markets.ziniao_matches("ML", "falabella-智利")
    assert not markets.ziniao_matches("FA", "MercadoLibre-智利-本土")
    assert markets.ziniao_matches("FA", None)            # nothing said, nothing contradicted
    assert markets.platform_from_ziniao("falabella-智利") == "FA"
    assert markets.platform_from_ziniao("沃尔玛-墨西哥") is None


# ------------------------------------------------------------ legacy config

def test_legacy_config_is_unchanged(monkeypatch):
    use(monkeypatch, LEGACY)
    assert store_config.market_of("BOCINA_SM") == "ML-MX"
    assert store_config.market_of("IDG_PE") == "ML-PE"
    assert store_config.market_of("never-configured") == "ML-MX"
    assert store_config.stores_for_region("MX") == ["BOCINA_SM", "EWTTO_SM"]
    assert store_config.stores_for_market("ML-CL") == ["Tecnohub_CL"]
    assert store_config.stores_for_market("ML-PE") == ["IDG_PE"]
    assert store_config.regions() == ["MX", "CL", "PE"]
    assert store_config.markets_configured() == ["ML-MX", "ML-CL", "ML-PE"]
    assert store_config.batch_stores() == ["BOCINA_SM", "EWTTO_SM"]


# --------------------------------------------------------------- new config

def test_catalog_config(monkeypatch):
    use(monkeypatch, NEW)
    assert store_config.market_of("DAB_CL-Fala (Coco)") == "FA-CL"
    assert store_config.platform_of("Tecnohub_CL") == "ML"
    assert store_config.country_of("AIS_CL-Fala") == "CL"
    assert store_config.stores_for_market("FA-CL") == ["DAB_CL-Fala (Coco)"]
    assert store_config.stores_for_market("ML-CL") == ["Tecnohub_CL"]
    assert store_config.markets_configured() == ["ML-MX", "ML-CL", "FA-CL"]
    assert store_config.regions() == ["MX", "CL"]


def test_falabella_store_never_in_mercadolibre_lists(monkeypatch):
    d = use(monkeypatch, NEW)
    d["batches"] = {}                                     # fall back to the allowlist
    assert store_config.stores_for_region("CL") == ["Tecnohub_CL"]
    assert set(store_config.stores_for_market("FA-CL")) == {"DAB_CL-Fala (Coco)",
                                                           "AIS_CL-Fala"}
    assert "AIS_CL-Fala" not in store_config.batch_stores()


def test_batch_with_a_store_of_another_market_raises(monkeypatch):
    d = use(monkeypatch, NEW)
    d["batches"] = {"CL": ["Tecnohub_CL", "AIS_CL-Fala"]}  # legacy key = ML-CL
    with pytest.raises(store_config.StoreConfigError):
        store_config.stores_for_region("CL")


def test_catalog_requires_every_allowed_store(monkeypatch):
    d = use(monkeypatch, NEW)
    d["allowed"].append("Mystery_Store")
    with pytest.raises(store_config.StoreConfigError):
        store_config.market_of("Mystery_Store")


def test_bad_catalog_entries_raise(monkeypatch):
    d = use(monkeypatch, NEW)
    d["catalog"]["AIS_CL-Fala"] = {"platform": "FA", "country": "MX"}
    with pytest.raises(store_config.StoreConfigError):
        store_config.market_of("BOCINA_SM")
    d["catalog"]["AIS_CL-Fala"] = "CL"
    with pytest.raises(store_config.StoreConfigError):
        store_config.market_of("BOCINA_SM")


def test_require_platform_and_ziniao_check(monkeypatch):
    use(monkeypatch, NEW)
    assert store_config.require_platform(["DAB_CL-Fala (Coco)"], "FA")
    with pytest.raises(store_config.StoreConfigError):
        store_config.require_platform(["Tecnohub_CL", "AIS_CL-Fala"], "ML")
    assert store_config.check_ziniao_platform("AIS_CL-Fala", "falabella-智利")
    with pytest.raises(store_config.StoreConfigError):
        store_config.check_ziniao_platform("Tecnohub_CL", "falabella-智利")


def test_mercadolibre_batch_refuses_falabella_store(monkeypatch):
    use(monkeypatch, NEW)

    class A:
        stores = ["AIS_CL-Fala"]
        limit = None
    with pytest.raises(store_config.StoreConfigError):
        run_batch.resolve_stores(A())


def test_meli_site_reads_catalog_country(monkeypatch):
    use(monkeypatch, NEW)
    assert meli_site.of_store("Tecnohub_CL") == "cl"
    assert meli_site.of_store("BOCINA_SM") == "mx"
