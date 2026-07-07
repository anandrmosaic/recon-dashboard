import os
import json
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES_SHEETS = ['https://www.googleapis.com/auth/spreadsheets']
SCOPES_GMAIL  = ['https://www.googleapis.com/auth/gmail.send']
SCOPES_ALL    = SCOPES_SHEETS + SCOPES_GMAIL


def get_sheets_credentials(credentials_file=None, token_file=None):
    # Cloud: use Service Account JSON stored in GOOGLE_SA_KEY env var
    sa_key = os.environ.get('GOOGLE_SA_KEY')
    if sa_key:
        info = json.loads(sa_key)
        return service_account.Credentials.from_service_account_info(
            info, scopes=['https://www.googleapis.com/auth/spreadsheets']
        )
    # Local dev fallback: use OAuth token file
    return _oauth_creds(credentials_file, token_file)


def get_gmail_credentials(credentials_file=None, token_file=None):
    # Cloud: use OAuth token JSON stored in GMAIL_TOKEN_JSON env var
    token_json = os.environ.get('GMAIL_TOKEN_JSON')
    if token_json:
        info = json.loads(token_json)
        creds = Credentials.from_authorized_user_info(info, SCOPES_GMAIL)
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
        return creds
    # Local dev fallback: use token file
    return _oauth_creds(credentials_file, token_file)


def _oauth_creds(credentials_file, token_file):
    creds = None
    if token_file and os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES_ALL)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_file, SCOPES_ALL)
            creds = flow.run_local_server(port=8080)
        if token_file:
            with open(token_file, 'w') as f:
                f.write(creds.to_json())
    return creds


# Legacy wrapper so any older call sites still work
def get_credentials(credentials_file, token_file):
    return _oauth_creds(credentials_file, token_file)
