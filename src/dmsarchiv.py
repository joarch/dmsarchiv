import configparser
import json
import os
import shutil
import sys
from datetime import datetime, timedelta
from decimal import Decimal
from getopt import getopt, GetoptError
from typing import List, Dict

import requests
from requests.auth import HTTPBasicAuth

from common import _json_load
from export_excel import export_nach_excel

DEFAULT_PARAMETER_SECTION = "config.ini:PARAMETER"
DEFAULT_EXPORT_PARAMETER_SECTION = "config.ini:EXPORT"

PARAM_URL = "dms_api_url"
PARAM_USER = "dms_api_benutzer"
PARAM_PASSWD = "dms_api_passwort"

DEFAULT_EXPORT_VON_DATUM = "01.01.2010"

CLASSIFY_ATTRIBUTES_FILENAME = "classify_attributes.json"
FOLDERS_FILENAME = "folders.json"
TYPES_FILENAME = "types.json"


def export(profil=DEFAULT_PARAMETER_SECTION, export_profil=DEFAULT_EXPORT_PARAMETER_SECTION, export_von_datum=None,
           export_bis_datum=None, max_documents=None, tage_offset=None, debug=None):
    # TODO LOG File schreiben
    # TODO timeit Zeit loggen bzw. als info_dauer in ini speichern

    # DMS API Connect
    api_url, cookies = _connect(profil)

    # DMS API Connect Info
    api_statistics = _get_statistics(api_url, cookies)
    export_info = dict()
    export_info["info_api_download_count"] = api_statistics["uploadCount"]
    export_info["info_api_upload_count"] = api_statistics["downloadCount"]
    export_info["info_api_max_count"] = api_statistics["maxCount"]

    # DMS API Klassifizierungsattribute auslesen, wenn noch nicht vorhanden
    if not os.path.exists(CLASSIFY_ATTRIBUTES_FILENAME):
        classify_attributes = _get_classify_attributes(api_url, cookies)
        with open(CLASSIFY_ATTRIBUTES_FILENAME, 'w', encoding='utf-8') as outfile:
            json.dump(classify_attributes, outfile, ensure_ascii=False, indent=2, sort_keys=True, default=json_serial)

    # DMS API Order auslesen, wenn noch nicht vorhanden
    if not os.path.exists(FOLDERS_FILENAME):
        folders = _get_folders(api_url, cookies)
        with open(FOLDERS_FILENAME, 'w', encoding='utf-8') as outfile:
            json.dump(folders, outfile, ensure_ascii=False, indent=2, sort_keys=True, default=json_serial)

    # DMS API Dokumentenart auslesen, wenn noch nicht vorhanden
    if not os.path.exists(TYPES_FILENAME):
        types = _get_types(api_url, cookies)
        with open(TYPES_FILENAME, 'w', encoding='utf-8') as outfile:
            json.dump(types, outfile, ensure_ascii=False, indent=2, sort_keys=True, default=json_serial)

    # Konfiguration lesen
    parameter_export = _get_config(export_profil)

    export_von_datum = parameter_export["export_von_datum"] if export_von_datum is None else export_von_datum
    export_bis_datum = parameter_export["export_bis_datum"] if export_bis_datum is None else export_bis_datum
    max_documents = int(parameter_export["max_documents"]) if max_documents is None else max_documents
    tage_offset = int(parameter_export["tage_offset"]) if tage_offset is None else tage_offset
    export_parameter = _json_load(parameter_export["export_parameter_datei"])

    if debug is None:
        debug = parameter_export.get("debug") == "true"

    if not export_von_datum:
        export_von_datum = DEFAULT_EXPORT_VON_DATUM

    # DMS API Search
    export_info["info_letzter_export"] = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    export_info["info_letzter_export_von_datum"] = export_von_datum
    documents = _search_documents(api_url, cookies, export_von_datum, export_parameter.get("suchparameter_list"),
                                  bis_datum=export_bis_datum, max_documents=max_documents, debug=debug)

    # Dokumenten Export Informationen auswerten
    ctimestamps = list(map(lambda d: datetime.strptime(d["classifyAttributes"]["ctimestamp"], "%Y-%m-%d %H:%M:%S"),
                           documents))
    ctimestamps.sort()
    if len(ctimestamps) > 0:
        min_ctimestamp = ctimestamps[0]
        max_ctimestamp = ctimestamps[-1]
    else:
        min_ctimestamp = None
        max_ctimestamp = None

    if export_bis_datum and len(documents) == 0:
        raise RuntimeError("Achtung es wurden keine Dokumente exportiert. Bitte das Such 'bis_datum' erweitern.")

    if len(documents) >= max_documents:
        raise RuntimeError(f"Achtung es wurden evtl. nicht alle Dokumente exportiert, Anzahl >= {max_documents}. "
                           f"Das Such-Datum muss weiter eingeschränkt werden. "
                           f"Es wurde gesucht mit {export_von_datum} - {export_bis_datum}.")
    if export_bis_datum:
        # es gab eine Einschränkung bis Datum
        export_von_datum = max_ctimestamp.strftime("%d.%m.%Y")
        # - nächste Zeitscheibe in Export-Info schreiben
        export_bis_datum = datetime.strptime(export_bis_datum, "%d.%m.%Y") + timedelta(days=tage_offset)
        if export_bis_datum < datetime.now():
            export_bis_datum = export_bis_datum.strftime("%d.%m.%Y")
        else:
            # Ende erreicht der nächste Export läuft ohne bis Datum
            export_bis_datum = ""
    else:
        export_von_datum = datetime.now().strftime("%d.%m.%Y")

    export_info["info_letzter_export_anzahl_dokumente"] = len(documents)
    export_info["info_min_ctimestamp"] = min_ctimestamp.strftime("%d.%m.%Y") if min_ctimestamp else export_von_datum
    export_info["info_max_ctimestamp"] = max_ctimestamp.strftime("%d.%m.%Y") if max_ctimestamp else ""
    # - Export Parameter für den nächsten Export
    export_info["export_von_datum"] = export_von_datum
    export_info["export_bis_datum"] = export_bis_datum
    export_info["max_documents"] = max_documents
    export_info["tage_offset"] = tage_offset

    # DMS API Disconnect
    _disconnect(api_url, cookies)

    # Dokumente als JSON Datei speichern
    result = {
        "export_time": datetime.now().strftime("%d.%m.%Y %H:%M:%S"),
        "documents": documents}
    json_export_datei = export_parameter["json_export_datei"]
    json_export_datei_tmp = json_export_datei + "_tmp"
    with open(json_export_datei_tmp, 'w', encoding='utf-8') as outfile:
        json.dump(result, outfile, ensure_ascii=False, indent=2, sort_keys=True, default=json_serial)
    result["anzahl_exportiert"] = len(documents)
    anzahl_neu = len(documents)
    # neue und vorhandene Export Ergebnisse zusammenführen, falls vorhanden
    if os.path.exists(json_export_datei):
        with open(json_export_datei, encoding="utf-8") as file:
            result_vorher = json.load(file)
        doc_ids_new = [document["docId"] for document in result["documents"]]
        for document in result_vorher["documents"]:
            if document["docId"] not in doc_ids_new:
                result["documents"].append(document)
            else:
                anzahl_neu -= 1
    result["anzahl"] = len(result["documents"])
    result["anzahl_neu"] = anzahl_neu

    # Sortierung nach DocId
    result["documents"].sort(key=lambda document: document["docId"])

    # Speichern in JSON Datei und löschen temp. Export Datei
    with open(json_export_datei, 'w', encoding='utf-8') as outfile:
        json.dump(result, outfile, ensure_ascii=False, indent=2, sort_keys=True, default=json_serial)
    os.remove(json_export_datei_tmp)

    print(
        f"Dokumente geladen im Zeitraum: {export_von_datum} - {export_bis_datum}, "
        f"Anzahl geladen: {result['anzahl_exportiert']}, "
        f"Anzahl neu: {result['anzahl_neu']}, "
        f"Anzahl gesamt: {result['anzahl']}.")

    # wenn alle Dokumente bis zum aktuell Tag exportiert wurden,
    # wird die Excel Datei geschrieben und die JSON Datei als Temp.-Datei umbenannt
    if export_von_datum == datetime.now().strftime("%d.%m.%Y"):
        # Excel Export
        if export_parameter["export"]["export_format"] == "xlsx":
            export_nach_excel(result, export_parameter["export"])
        else:
            raise RuntimeError(f"nicht unterstütztes Export Format {export_parameter['export']['export_format']}")
        # vorhandene JSON Datei als Temp.-Datei sichern
        splitext = os.path.splitext(json_export_datei)
        shutil.move(json_export_datei, os.path.join(
            os.path.dirname(json_export_datei),
            os.path.basename(splitext[0]) + "_tmp" + splitext[1]
        ))
    else:
        # noch nicht alle Dokumente geladen
        print("Es wurden noch nicht alle Dokumente bis zum heutigen Tag geladen, der Export wird nicht durchgeführt.")
        print("Bitte das Programm erneut ausführen.")

    # Export Info (letzter Export Zeitstempel und DMS API Info) in die Config-Datei zurückschreiben
    _write_config(export_profil, export_info)


def _search_documents(api_url, cookies, von_datum, suchparameter_list=None,
                      bis_datum=None, max_documents=1000, debug=False) -> List[Dict]:
    suchparameter_list = suchparameter_list or []
    von_datum = datetime.strptime(von_datum, "%d.%m.%Y")
    # Search-Date -1 Tag, vom letzten Lauf aus,
    # da die DMS API Suche nicht mit einem Zeitstempel umgehen kann
    # zusätzlich (sicherheitshalber) Vergleich mit >=
    # von_datum = von_datum.date() - timedelta(days=1)
    von_datum = von_datum.date()

    von_datum = von_datum.strftime("%Y-%m-%d")
    search_parameter = [{"classifyAttribut": "ctimestamp", "searchOperator": ">=",
                         "searchValue": von_datum}]
    if bis_datum:
        bis_datum = datetime.strptime(bis_datum, "%d.%m.%Y").strftime("%Y-%m-%d")
        search_parameter.append({"classifyAttribut": "ctimestamp", "searchOperator": "<=",
                                 "searchValue": bis_datum})

    for suchparameter in suchparameter_list:
        search_parameter.append(suchparameter)

    such_data = json.dumps(search_parameter)

    if debug:
        print(f"Suche mit: {json.dumps(search_parameter)}")

    r = requests.post("{}/searchDocumentsExt?maxDocumentCount={}".format(api_url, max_documents),
                      data=such_data,
                      cookies=cookies, headers=_headers())
    _assert_request(r)
    documents = json.loads(r.text)

    if debug:
        print(f"Suche Fertig. Anzahl Dokumente : {len(documents)}")

    return documents


def _get_statistics(api_url, cookies):
    r = requests.get("{}/apiStatistics".format(api_url), cookies=cookies, headers=_headers())
    _assert_request(r)
    return json.loads(r.text)


def _get_classify_attributes(api_url, cookies):
    r = requests.get("{}/classifyAttributes".format(api_url), cookies=cookies, headers=_headers())
    _assert_request(r)
    return json.loads(r.text)


def _get_folders(api_url, cookies):
    r = requests.get("{}/folders".format(api_url), cookies=cookies, headers=_headers())
    _assert_request(r)
    return json.loads(r.text)


def _get_types(api_url, cookies):
    r = requests.get("{}/types".format(api_url), cookies=cookies, headers=_headers())
    _assert_request(r)
    return json.loads(r.text)


def _headers():
    return {'Content-Type': 'application/json; charset=utf8'}


def _connect(profil):
    params = _get_config(profil)
    r = requests.get("{}/connect/1".format(params[PARAM_URL]),
                     auth=HTTPBasicAuth(params[PARAM_USER], params[PARAM_PASSWD]))

    _assert_request(r)

    cookies = r.cookies.get_dict()

    return params[PARAM_URL], cookies


def _disconnect(api_url, cookies):
    r = requests.get("{}/disconnect".format(api_url), cookies=cookies)

    _assert_request(r)


def _assert_request(request):
    if request.status_code != 200:
        raise RuntimeError(f"Fehler beim Request: {request.status_code}, Message: {request.text}")


def _get_config(profil):
    split = profil.split(":")
    config_file = split[0]
    config_section = split[1]
    config = configparser.ConfigParser()
    config.read(config_file)
    return config[config_section]


def _write_config(profil, new_params):
    """
    Aktualisiert die Config-Datei mit neuen Werten.
    """
    split = profil.split(":")
    config_file = split[0]
    config_section = split[1]
    config = configparser.ConfigParser()
    config.read(config_file)
    # merge alte und neue Parameter
    for section in config.sections():
        if section == config_section:
            for key, value in new_params.items():
                config[section][key] = str(value)

    with open(config_file, 'w') as configfile:
        config.write(configfile)


def json_serial(obj):
    if isinstance(obj, datetime):
        serial = obj.isoformat()
        return serial
    if isinstance(obj, Decimal):
        serial = str(obj)
        return serial
    raise TypeError("Type not serializable")


def main(argv):
    """
    Export DMS Dokumenten Infos. Das Zielformat wird über das Export Profil übergeben.
    Programmargumente:
    - parameter (INI-Datei und Section): z.B.: 'config.ini:PARAMETER'
    - export_parameter (INI-Datei und Section): z.B.: 'config.ini:EXPORT'
    """
    hilfe = f"{os.path.basename(__file__)} -p <parameter> -e <export_parameter>"
    parameter = ""
    export_parameter = ""
    try:
        opts, args = getopt(argv, "hp:e:", ["parameter=", "export_parameter="])
    except GetoptError:
        print(hilfe)
        sys.exit(2)

    for opt, arg in opts:
        if opt in ("-h", "--help"):
            print(hilfe)
            sys.exit()
        elif opt in ("-p", "--parameter"):
            parameter = arg
        elif opt in ("-e", "--export_parameter"):
            export_parameter = arg

    if not parameter:
        parameter = DEFAULT_PARAMETER_SECTION
    if not export_parameter:
        export_parameter = DEFAULT_EXPORT_PARAMETER_SECTION

    export(parameter, export_parameter)


if __name__ == '__main__':
    main(sys.argv[1:])
