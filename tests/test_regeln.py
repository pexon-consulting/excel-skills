import datetime as dt

import pytest

from excel_skills import regeln


@pytest.mark.parametrize("text, wert", [
    ("1.234,56", 1234.56),
    ("-1.234,56", -1234.56),
    ("1 234,56", 1234.56),
    ("1\u00a0234,56", 1234.56),
    ("12,5", 12.5),
    ("12.500,00", 12500.0),
])
def test_deutsche_zahl_wird_zahl(text, wert):
    assert regeln.deutsche_zahl(text).wert == wert


@pytest.mark.parametrize("text", ["1.234", "1.234.567"])
def test_punkt_ohne_komma_ist_mehrdeutig(text):
    z = regeln.deutsche_zahl(text)
    assert z.wert is None and z.mehrdeutig


@pytest.mark.parametrize("text", ["01067", "4711", "K-1001", "1,2,3", "abc"])
def test_kein_zahltext(text):
    z = regeln.deutsche_zahl(text)
    assert z.wert is None and not z.mehrdeutig


def test_datum():
    assert regeln.deutsches_datum("31.12.2026") == dt.datetime(2026, 12, 31)
    assert regeln.deutsches_datum("30.04.26") == dt.datetime(2026, 4, 30)
    assert regeln.deutsches_datum("01.01.30") == dt.datetime(1930, 1, 1)
    assert regeln.deutsches_datum("31.02.2026") is None
    assert regeln.deutsches_datum("12/31/2026") is None
    assert regeln.ist_us_datum("12/31/2026")


@pytest.mark.parametrize("werte, erwartet", [
    (["Summe", None, 10], True),
    ([None, "Zwischensumme Nord", 5], True),
    (["gesamt", 3], True),
    (["Total"], True),
    (["Nord", "Summe"], False),
    ([12, "Kunde"], False),
    (["=SUM(A1:A3)", "Summe"], True),
])
def test_summenzeile(werte, erwartet):
    assert regeln.ist_summenzeile(werte) is erwartet


def test_kopfzeile_mit_titel_darueber():
    zeilen = [
        [("Bericht", False), (None, False), (None, False)],
        [(None, False)] * 3,
        [("Kunde", False), ("Ort", False), ("Umsatz", False)],
        [("Muster GmbH", False), ("Kiel", False), (10, False)],
    ]
    idx, grund = regeln.finde_kopfzeile(zeilen, 3)
    assert idx == 2 and "Zahl" in grund


def test_kopfzeile_nur_text_nimmt_erste():
    zeilen = [[("Name", False), ("Ort", False), ("Mail", False)], [("A", False), ("B", False), ("C", False)]]
    assert regeln.finde_kopfzeile(zeilen, 3)[0] == 0


def test_schema_enthaelt_alle_regeln():
    s = regeln.schema()
    for k in ("kopfzeile", "tabelle_oder_bericht", "deutsche_zahl", "datum", "summenzeile", "formelspalte", "csv_formate"):
        assert k in s
