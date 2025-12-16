
import requests

TENANT_ID = ""
CLIENT_ID = ""
CLIENT_SECRET = ""
SP_HOST   = ""
SITE_PATH = ""
PERIOD = "D90"


# token
tok = requests.post(
    f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token",
    data={"grant_type":"client_credentials","client_id":CLIENT_ID,
          "client_secret":CLIENT_SECRET,"scope":"https://graph.microsoft.com/.default"}
)
tok.raise_for_status()
AT = tok.json()["access_token"]
H  = {"Authorization": f"Bearer {AT}"}

# Get the site by path
resp = requests.get(
    "https://graph.microsoft.com/v1.0/sites/",
    headers=H
)
resp.raise_for_status()
j = resp.json()
site_id = j["id"]          # Graph composite id ("{hostname},{siteGuid},{webGuid}")
print("Graph site id:", site_id, "  webUrl:", j["webUrl"])

# The site collection GUID for the report is the middle part:
site_guid = site_id.split(",")[1]
print("Site GUID (matches 'Site Id' in report):", site_guid)
