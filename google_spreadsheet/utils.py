# encoding=utf8
'''
Created on 2016-11-10

@author: jingyang <jingyang@nexa-corp.com>
'''

import os

from oauth2client.client import GoogleCredentials
from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient.discovery import build


BASE_DIR = os.path.dirname(os.path.dirname(__file__))


def spreadsheet_service(on_gce=False, key_file_location=None, scopes=None):
    """
    Get client service to request spreadsheet APIs

    :param on_gce: project runs on Google Compute Engine or not
    :param key_file_location: json-formatted API key file
    :param scopes: OAuth 2.0 scopes, doc: https://developers.google.com/sheets/guides/authorizing#OAuth2Authorizing
    :return: client service object
    """
    if scopes is None:
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
        ]
    if on_gce:
        credentials = GoogleCredentials.get_application_default()
    else:
        credentials = ServiceAccountCredentials.from_json_keyfile_name(key_file_location, scopes)
    return build("sheets", "v4", credentials=credentials)
