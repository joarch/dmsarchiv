import json
from json.decoder import JSONDecodeError


def _json_load(filename):
    encodings = ['utf-8-sig', 'utf-8', 'windows-1250', 'windows-1252', 'iso-8859-1', 'cp1252']
    for encoding in encodings:
        try:
            with open(filename, encoding=encoding) as file:
                export_parameter = json.load(file)
            return export_parameter
        except JSONDecodeError:
            # TODO log warning
            pass
    raise RuntimeError(
        f"Fehler beim Lesen der Datei '{filename}', unbekanntes Encoding Format. "
        f"Folgende Formate wurden nicht erkannt '{encodings}'.")
