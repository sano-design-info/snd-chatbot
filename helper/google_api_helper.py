import os
from pathlib import Path
from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow


# load config
load_dotenv()

# generate Path
parent_dirpath = Path(__file__).parents[1]
cred_filepath = os.environ.get("CRED_FILEPATH")
token_save_path = parent_dirpath / "token.json"
cred_json = parent_dirpath / cred_filepath


def get_cledential(scopes: list[str]) -> Credentials:
    creds = None
    if token_save_path.exists():
        creds = Credentials.from_authorized_user_file(token_save_path, scopes)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(cred_json, scopes)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with token_save_path.open("w") as token:
            token.write(creds.to_json())
    return creds
