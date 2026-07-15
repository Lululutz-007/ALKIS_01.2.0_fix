#!/usr/bin/env python3
"""Automatischer Freigabetest für ALKIS_GeoAS_GEOgraf V1.2.0."""
from __future__ import annotations

import importlib.util
import tempfile
from pathlib import Path

from lxml import etree
from openpyxl import Workbook


ROOT = Path(__file__).resolve().parent
APP_PATH = ROOT / "ALKIS_GeoAS_GEOgraf.py"


def load_app():
    spec = importlib.util.spec_from_file_location("alkis_app", APP_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError("Quellmodul konnte nicht geladen werden.")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def assert_true(condition, message):
    if not condition:
        raise AssertionError(message)


def make_nas(app, path, parcel_ids):
    nsmap = {"wfs": app.W, "adv": app.A, "gml": app.G, "xlink": app.X}
    fc = etree.Element(f"{{{app.W}}}FeatureCollection", nsmap=nsmap)
    for parcel_id in parcel_ids:
        member = etree.SubElement(fc, f"{{{app.W}}}member")
        parcel = etree.SubElement(member, f"{{{app.A}}}AX_Flurstueck")
        parcel.set(f"{{{app.G}}}id", parcel_id)
        etree.SubElement(
            parcel, f"{{{app.A}}}flurstueckskennzeichen"
        ).text = parcel_id[-4:]
        etree.SubElement(parcel, f"{{{app.A}}}weistAuf")
    etree.ElementTree(fc).write(path, encoding="UTF-8", xml_declaration=True)


def make_geoas(path, parcel_ids):
    headers = [
        "ALKIS_ObjektID", "Grundbuchblatt", "Buchungsart",
        "Laufende Nummer", "Namensnummer", "Nachname oder Firma",
        "Grundbuchbezirk-Schlüssel", "Grundbuchbezirk-Name",
        "Grundbuchamt-Schlüssel", "Grundbuchamt-Name",
        "Grundbuchblattnummer", "Buchungsblattart",
        "Miteigentumsanteil", "Nummer im Aufteilungsplan",
        "Art der Rechtsgemeinschaft", "Beschrieb der Rechtsgemeinschaft",
        "Gemeinschaftl. Anteil", "Vorname", "geb.", "Geburtsdatum",
        "Akademischer Grad", "Anrede-Schlüssel", "PLZ", "Ort",
        "Straße", "Hausnummer", "PLZ Postfach", "Postfach", "Land",
    ]
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(headers)

    # Zwei Flurstücke gehören zu demselben 1302-Komplex. Die beiden
    # Buchungsstellen besitzen Dezimalanteile und müssen zusammen 1 ergeben.
    for parcel_id in parcel_ids:
        for laufende_nummer, anteil in (("1", "0,5/1"), ("2", "0.5/1")):
            row = {header: "" for header in headers}
            row.update({
                "ALKIS_ObjektID": parcel_id,
                "Grundbuchblatt": "0312340000001",
                "Buchungsart": "1302",
                "Laufende Nummer": laufende_nummer,
                "Grundbuchbezirk-Schlüssel": "031234",
                "Grundbuchbezirk-Name": "Testbezirk",
                "Grundbuchamt-Schlüssel": "0312",
                "Grundbuchamt-Name": "Testamt",
                "Grundbuchblattnummer": "1",
                "Buchungsblattart": "1000",
                "Miteigentumsanteil": anteil,
            })
            sheet.append([row[header] for header in headers])

    workbook.save(path)


def main():
    app = load_app()
    assert_true(app.VERSION == "1.2.0", "Versionskonstante ist nicht 1.2.0.")

    comma_fraction = app.frac("129,5/1000")
    point_fraction = app.frac("0.5/1")
    assert_true(comma_fraction is not None, "Komma-Dezimalanteil wird nicht erkannt.")
    assert_true(point_fraction is not None, "Punkt-Dezimalanteil wird nicht erkannt.")
    assert_true(
        app.decimal_text(comma_fraction[0]) == "129.5",
        "Komma-Dezimalanteil wird nicht korrekt normalisiert.",
    )

    with tempfile.TemporaryDirectory(prefix="alkis_v120_test_") as tmp:
        tmp = Path(tmp)
        parcel_ids = ["DENIAL43TESTPARC01", "DENIAL43TESTPARC02"]
        nas = tmp / "nas.xml"
        xlsx = tmp / "geoas.xlsx"
        output = tmp / "output.xml"
        report = tmp / "report.txt"
        starter = tmp / "starter.html"

        make_nas(app, nas, parcel_ids)
        make_geoas(xlsx, parcel_ids)
        app.generate_nas(nas, xlsx, output, report)

        # Windows-Regressionsprüfung: Die GeoAS-Datei muss nach dem Lesen
        # wieder freigegeben sein. Ein offenes openpyxl-Handle verursacht
        # sonst WinError 32 beim Umbenennen oder beim Aufräumen des Temp-Ordners.
        unlocked_xlsx = tmp / "geoas_unlocked.xlsx"
        xlsx.rename(unlocked_xlsx)
        unlocked_xlsx.rename(xlsx)

        app.write_html(starter, 'https://example.invalid/?x=1&y="z"')

        tree = etree.parse(str(output))
        bookings = tree.xpath("//adv:AX_Buchungsstelle", namespaces=app.NS)
        codes = [
            booking.xpath("string(adv:buchungsart)", namespaces=app.NS)
            for booking in bookings
        ]

        assert_true(codes.count("1102") == 1, "Es wurde nicht genau eine 1102-Wurzel erzeugt.")
        assert_true(codes.count("1302") == 2, "Es wurden nicht genau zwei 1302-Stellen erzeugt.")

        parcel_links = [
            tree.xpath(
                "string(//adv:AX_Flurstueck[@gml:id=$parcel_id]"
                "/adv:istGebucht/@xlink:href)",
                namespaces=app.NS,
                parcel_id=parcel_id,
            )
            for parcel_id in parcel_ids
        ]
        assert_true(
            len(set(parcel_links)) == 1 and parcel_links[0],
            "Die Flurstücke verweisen nicht auf dieselbe 1102-Wurzel.",
        )

        report_text = report.read_text(encoding="utf-8")
        assert_true("Version: 1.2.0" in report_text, "Prüfprotokoll nennt nicht Version 1.2.0.")
        assert_true("Komplex_1302: 1" in report_text, "1302-Komplexstatistik fehlt.")
        assert_true("Dezimalanteile: 2" in report_text, "Dezimalstatistik fehlt.")
        assert_true("Doppelte gml:id: 0" in report_text, "Doppelte IDs im Test.")
        assert_true(
            "Gebrochene Referenzen in erzeugten Ketten: 0" in report_text,
            "Gebrochene Referenzen im Test.",
        )
        assert_true("Warnungen: 0" in report_text, "Der Integrationstest erzeugt Warnungen.")

        html_text = starter.read_text(encoding="utf-8")
        assert_true("<!doctype html>" in html_text.lower(), "HTML-Starter fehlt.")
        assert_true("window.location.replace" in html_text, "HTML-Weiterleitung fehlt.")
        assert_true("&amp;" in html_text, "URL wird im HTML nicht sicher maskiert.")

    print("V1.2.0-Freigabetest: ALLE PRÜFUNGEN BESTANDEN")


if __name__ == "__main__":
    main()
