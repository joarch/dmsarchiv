# -*- coding: utf-8 -*-

import json
import os
import re
import sys
from datetime import datetime, date
from decimal import Decimal
from getopt import getopt, GetoptError

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter


def export_nach_excel(documents, export_profil):
    wb = Workbook()
    ws = wb.active
    # ws.title = ...

    row = 1

    # mit Spaltenüberschrifen
    if export_profil["spaltenueberschrift"].lower() == "ja":
        column_header_format = export_profil.get("spaltenueberschrift_format")
        if column_header_format is not None:
            if "PatternFill" == column_header_format["format"]:
                column_header = PatternFill(start_color=column_header_format["start_color"],
                                            end_color=column_header_format["end_color"],
                                            fill_type=column_header_format["fill_type"])
            else:
                raise RuntimeError(
                    f"Unbekanntes Format {column_header_format['format']} in 'spaltenueberschrift_format/format'. "
                    f"Möglich ist nur 'PatternFill'")
        else:
            # Standard Format
            column_header = PatternFill(start_color='AAAAAA',
                                        end_color='AAAAAA',
                                        fill_type='solid')
        column = 1
        for spalte in export_profil["spalten"]:
            ws.cell(column=column, row=row, value=spalte["ueberschrift"])
            col = ws["{}{}".format(get_column_letter(column), row)]
            col.font = Font(bold=True)
            col.fill = column_header
            column += 1
        row += 1

    for document in documents["documents"]:
        column = 1

        for spalte in export_profil["spalten"]:
            feld_name = spalte["feld"]
            mapped_value = ""

            if feld_name:
                # Spalten Wert auslesen und mappen
                if feld_name in document:
                    value = document[feld_name]
                elif feld_name in document["classifyAttributes"]:
                    value = document["classifyAttributes"][feld_name]
                else:
                    raise RuntimeError(
                        f"Die Spalte '{feld_name}' existiert nicht im Dokument. Bitte Export-Profil überprüfen.")
                mapped_value = map_str_value(value)
            else:
                # keine Feld Name, damit bleibt die Spalte leer
                pass

            cell = ws.cell(column=column, row=row, value=mapped_value)
            if spalte.get("number_format"):
                cell.number_format = spalte["number_format"]
            else:
                if isinstance(mapped_value, date):
                    cell.number_format = 'DD.MM.YYYY'
                if isinstance(mapped_value, datetime):
                    cell.number_format = 'DD.MM.YYYY HH:MM:SS'
            column += 1
        row += 1

    wb.save(filename=export_profil["dateiname"])


def map_str_value(value):
    if type(value) != str:
        return value

    if value == "undefined":
        # clean up
        value = ""

    if value == "true":
        value = "ja"

    if value == "false":
        value = "nein"

    if "€" in value \
            and (value[0].isnumeric() or len(value) >= 2 and value[0] == "-" and value[1].isnumeric()):
        return map_eur(value)

    eur_pattern = re.compile(r"^-?[0-9]+,?[0-9]* (€|EUR)$")
    if eur_pattern.match(value):
        return map_eur(value)

    datum_pattern = re.compile(r"^[0-9]{2}\.[0-9]{2}\.[0-9]{4}$")
    if datum_pattern.match(value):
        return map_datum(value)

    datum_pattern = re.compile(r"^[0-9]{4}-[0-9]{2}-[0-9]{2}$")
    if datum_pattern.match(value):
        return map_datum(value)

    datum_pattern = re.compile(r"^[0-9]{2}\.[0-9]{2}\.[0-9]{4} [0-9]{2}:[0-9]{2}:[0-9]{2}$")
    if datum_pattern.match(value):
        return map_datum_zeit(value)

    datum_pattern = re.compile(r"^[0-9]{4}-[0-9]{2}-[0-9]{2} [0-9]{2}:[0-9]{2}:[0-9]{2}$")
    if datum_pattern.match(value):
        return map_datum_zeit(value)

    decimal_pattern = re.compile(r"^-?[0-9]+,?[0-9]*$")
    if decimal_pattern.match(value):
        return map_number(value)

    return value


def map_number(value):
    if value is None:
        return None
    return Decimal(value.replace('.', '').replace(' ', '').replace(',', '.'))


def map_eur(value):
    return map_number(value.replace("€", "").replace("EUR", ""))


def map_datum(value):
    if "-" in value:
        return datetime.strptime(value, "%Y-%m-%d").date()
    return datetime.strptime(value, "%d.%m.%Y").date()


def map_datum_zeit(value):
    if "-" in value:
        return datetime.strptime(value, "%Y-%m-%d %H:%M:%S")
    return datetime.strptime(value, "%d.%m.%Y %H:%M:%S")


def main(argv):
    """
    Export die übergebene JSON Datei (documents_datei) mit den exportierten DMS Dokumenten Feldern nach Excel.
    Das Export Format wird mit der übergebenen Export Parameter Datei (export_parameter_datei) konfiguriert.
    """
    hilfe = f"{os.path.basename(__file__)} -d <documents_datei> -e <export_parameter_datei>"
    documents_datei = ""
    export_parameter_datei = ""
    try:
        opts, args = getopt(argv, "hd:e:", ["documents_datei=", "export_parameter_datei="])
    except GetoptError:
        print(hilfe)
        sys.exit(2)

    for opt, arg in opts:
        if opt in ("-h", "--help"):
            print(hilfe)
            sys.exit()
        elif opt in ("-d", "--documents_datei"):
            documents_datei = arg
        elif opt in ("-e", "--export_parameter_datei"):
            export_parameter_datei = arg

    if not documents_datei or not export_parameter_datei:
        print("Usage: " + hilfe)
        sys.exit(2)

    if not os.path.exists(documents_datei):
        raise RuntimeError(f"Die Datei '{documents_datei}' existiert nicht.")
    if not os.path.exists(export_parameter_datei):
        raise RuntimeError(f"Die Datei '{export_parameter_datei}' existiert nicht.")

    with open(documents_datei, encoding="utf-8") as file:
        documents = json.load(file)
    with open(export_parameter_datei, encoding="utf-8") as file:
        export_parameter = json.load(file)

    export_nach_excel(documents, export_parameter["export"])


if __name__ == '__main__':
    main(sys.argv[1:])
