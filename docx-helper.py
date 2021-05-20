import httplib2
import pygsheets

from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient import discovery
from googleapiclient import errors
from pydrive.auth import GoogleAuth

cred_json_file = '../conf/credential.json'
gsheet_name = 'NBR-AW-NSW__comparison-study'
_G = pygsheets.authorize(service_account_file=cred_json_file)

credentials = ServiceAccountCredentials.from_json_keyfile_name(cred_json_file, scopes=['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets'])
credentials.authorize(httplib2.Http())

gauth = GoogleAuth()
gauth.credentials = credentials

_service = discovery.build('sheets', 'v4', credentials=credentials)

gsheet = _G.open(gsheet_name)



from pydrive.drive import GoogleDrive

def copy_file(service, origin_file_id, copy_title):
    """
    Copy an existing file.

    Args:
        service: Drive API service instance.
        origin_file_id: ID of the origin file to copy.
        copy_title: Title of the copy.

    Returns:
        The copied file if successful, None otherwise.
    """
    copied_file = {'name': copy_title}
    try:
        return service.files().copy(fileId=origin_file_id, body=copied_file).execute()
    except(errors.HttpError, error):
        print('An error occurred: {0}'.format(error))
        return None

GSHEET_LIST = ["LGRD-UPEHSDP__KHA__1__forms-technical",
    "LGRD-UPEHSDP__KHA__2A__organization-documents-spectrum",
    "LGRD-UPEHSDP__KHA__2B__organization-documents-bracnet",
    "LGRD-UPEHSDP__KHA__2C__organization-documents-bat",
    "LGRD-UPEHSDP__KHA__3A__financial-reports-spectrum",
    "LGRD-UPEHSDP__KHA__3B__financial-reports-bracnet",
    "LGRD-UPEHSDP__KHA__3C__financial-reports-bat",
    "LGRD-UPEHSDP__KHA__4A__contract-wo-wcc-spectrum",
    "LGRD-UPEHSDP__KHA__4B__contract-wo-wcc-bracnet",
    "LGRD-UPEHSDP__KHA__4C__contract-wo-wcc-bat",
    "LGRD-UPEHSDP__KHA__5__manufacturer's-authorization",
    "LGRD-UPEHSDP__KHA__6__product-data-sheet"
]

for gsheet_name in GSHEET_LIST:
    sheet = _G.open(gsheet_name)
    title = gsheet_name.replace("__KHA__", "__GHA__")
    copy_file(_service, sheet.id, title)

# GSHEETS.append({'name': gsheet_name, 'id': sheet.id})
GSHEETS = []

pprint.pprint(GSHEETS)
