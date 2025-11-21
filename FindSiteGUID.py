
import requests

TENANT_ID = "830138cc-a829-4041-926e-dea50549d68b"
CLIENT_ID = "e377d011-c16f-4a0d-8c69-14091ec04ff2"
CLIENT_SECRET = "HnJ8Q~nFySE58UbB2bsL6RvESZQ1VIF6Q5C~ObWv"
SP_HOST   = "https://tiongseng.sharepoint.com"
SITE_PATH = "sites/QS"
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
    "https://graph.microsoft.com/v1.0/sites/tiongseng.sharepoint.com:/sites/QS?$select=id,webUrl",
    headers=H
)
resp.raise_for_status()
j = resp.json()
site_id = j["id"]          # Graph composite id ("{hostname},{siteGuid},{webGuid}")
print("Graph site id:", site_id, "  webUrl:", j["webUrl"])

# The site collection GUID for the report is the middle part:
site_guid = site_id.split(",")[1]
print("Site GUID (matches 'Site Id' in report):", site_guid)