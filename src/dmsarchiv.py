import configparser
import json
import os
from datetime import datetime, timedelta
from decimal import Decimal
from shutil import copyfile

import requests
from requests.auth import HTTPBasicAuth

DEFAULT_PROFIL = "config/config.ini:PARAMETER"
DEFAULT_EXPORT_PROFIL = "config/config.ini:EXPORT"

PARAM_URL = "dms_api_url"
PARAM_USER = "dms_api_benutzer"
PARAM_PASSWD = "dms_api_passwort"

MIN_DATETIME = datetime.strptime("01.01.2000", "%d.%m.%Y")
MAX_DATETIME = datetime.strptime("01.01.3999", "%d.%m.%Y")


def export(profil=DEFAULT_PROFIL, export_profil=DEFAULT_EXPORT_PROFIL, bis_datum=None):
    # TODO timeit Zeit loggen bzw. als info_dauer in ini speichern

    # DMS API Connect
    api_url, cookies = _connect(profil)

    # DMS API Connect Info
    api_statistics = _get_statistics(api_url, cookies)
    export_info = dict()
    export_info["info_api_download_count"] = api_statistics["uploadCount"]
    export_info["info_api_upload_count"] = api_statistics["downloadCount"]
    export_info["info_api_max_count"] = api_statistics["maxCount"]

    # Konfiguration lesen
    parameter = _get_config(profil)
    export_von_datum = _get_config(export_profil)["export_von_datum"]

    # DMS API Search
    max_documents = int(parameter["max_documents"])
    export_info["info_letzter_export"] = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    export_info["info_letzter_export_von_datum"] = export_von_datum
    documents = _search_documents(api_url, cookies, export_von_datum, bis_datum=bis_datum, max_documents=max_documents)

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

    if bis_datum is not None and len(documents) == 0:
        raise RuntimeError("Achtung es wurden keine Dokumente exportiert. Bitte das Such 'bis_datum' erweitern.")

    if len(documents) >= max_documents:
        raise RuntimeError("Achtung es wurden evtl. nicht alle Dokumente exportiert, Anzahl >= {}."
                           " Das Such-Datum muss weiter eingeschränkt werden. min_ctimestamp={}, max_ctimestamp={}."
                           .format(max_documents, min_ctimestamp.strftime("%d.%m.%Y"),
                                   max_ctimestamp.strftime("%d.%m.%Y")))
    if bis_datum is not None:
        export_von_datum = max_ctimestamp.strftime("%d.%m.%Y")
    else:
        export_von_datum = datetime.now().strftime("%d.%m.%Y")

    export_info["info_letzter_export_anzahl_dokumente"] = len(documents)
    export_info["info_min_ctimestamp"] = min_ctimestamp.strftime("%d.%m.%Y")
    export_info["info_max_ctimestamp"] = max_ctimestamp.strftime("%d.%m.%Y")
    # - Export-Von-Datum für den nächsten Export
    export_info["export_von_datum"] = export_von_datum

    # DMS API Disconnect
    _disconnect(api_url, cookies)

    # Dokumente als JSON Datei speichern
    result = {
        "anzahl": len(documents),
        "export_time": datetime.now().strftime("%d.%m.%Y %H:%M:%S"),
        "documents": documents}
    json_export_datei = parameter["json_export_datei"]
    json_export_datei_tmp = json_export_datei + "_tmp"
    with open(json_export_datei_tmp, 'w', encoding='utf-8') as outfile:
        json.dump(result, outfile, ensure_ascii=False, indent=2, sort_keys=True, default=json_serial)
    copyfile(json_export_datei_tmp, json_export_datei)
    os.remove(json_export_datei_tmp)

    # TODO LOG File schreiben

    # TODO mergen in eine JSON Datei

    # TODO JSON als CSV oder Excel speichern

    # Export Info (letzter Export Zeitstempel und DMS API Info) in die Config-Datei zurückschreiben
    _write_config(export_profil, export_info)


def _search_documents(api_url, cookies, von_datum, bis_datum=None, max_documents=1000):
    von_datum = datetime.strptime(von_datum, "%d.%m.%Y")
    # Search-Date -1 Tag, vom letzten Lauf aus,
    # da die DMS API Suche nicht mit einem Zeitstempel umgehen kann
    # TODO sicherheitshalber oder reicht >=
    von_datum = von_datum.date() - timedelta(days=1)

    von_datum = von_datum.strftime("%Y-%m-%d")
    search_parameter = [{"classifyAttribut": "ctimestamp", "searchOperator": ">=",
                         "searchValue": von_datum}]
    if bis_datum is not None:
        bis_datum = datetime.strptime(bis_datum, "%d.%m.%Y").strftime("%Y-%m-%d")
        search_parameter.append({"classifyAttribut": "ctimestamp", "searchOperator": "<=",
                                 "searchValue": bis_datum})

    # TODO weitere Suchparameter aus config/dmsarchive.json verwenden (export_parameter)

    r = requests.post("{}/searchDocumentsExt?maxDocumentCount={}".format(api_url, max_documents),
                      data=json.dumps(search_parameter),
                      cookies=cookies, headers=_headers())
    _assert_request(r)
    return json.loads(r.text)


def _get_statistics(api_url, cookies):
    r = requests.get("{}/apiStatistics".format(api_url), cookies=cookies, headers=_headers())
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


if __name__ == '__main__':
    pass
