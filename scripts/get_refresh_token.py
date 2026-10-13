import os
from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive",
]

CLIENT_FILE = os.path.join(os.path.dirname(__file__), "..", "secrets", "oauth-client.json")

flow = InstalledAppFlow.from_client_secrets_file(CLIENT_FILE, SCOPES)
creds = flow.run_local_server(port=0, prompt="consent", access_type="offline")

print("=" * 60)
print("REFRESH TOKEN:")
print(creds.refresh_token)
print("=" * 60)
print("CLIENT ID:")
print(creds.client_id)
print("=" * 60)
print("CLIENT SECRET:")
print(creds.client_secret)
print("=" * 60)