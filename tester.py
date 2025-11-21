# pip install msal requests pandas openpyxl
import os, io, time, json, requests, msal, pandas as pd
from datetime import datetime, timedelta, timezone

# ========= OPTIONAL: Corp proxy/TLS =========
# os.environ["HTTPS_PROXY"] = "http://proxy:8080"
# os.environ["REQUESTS_CA_BUNDLE"] = r"C:\path\to\cacert.pem"
# os.environ["SSL_CERT_FILE"]     = r"C:\path\to\cacert.pem"
PROXIES = None   # or {"https": "http://proxy:8080"}
VERIFY  = True   # or r"C:\path\to\cacert.pem"
# ===========================================

# ==== inputs (change these when switching sites) ====
TENANT_ID = "830138cc-a829-4041-926e-dea50549d68b"
CLIENT_ID = "e377d011-c16f-4a0d-8c69-14091ec04ff2"
HOSTNAME  = "tiongseng.sharepoint.com"

# Read your client secret from env (recommended). If missing, raise an error.
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
if not CLIENT_SECRET:
    raise RuntimeError("Set CLIENT_SECRET environment variable to your app's client secret.")

# SITE PATHS & IDS
SITE_PATH = "sites/Project-SDEManagement"  # change only this line when switching sites
SITE_COMPOSITE_ID = "tiongseng.sharepoint.com,b6c68d09-2c2c-4a8c-8557-1217bd3506bb,1fe88523-7a1b-4cfb-be57-00d05b591233"

SITE_PAGES_TITLE = "Site Pages"   # change only if the library is named differently
DAYS_DAILY = 90                   # daily window
FALLBACK_TOTAL_DAYS = 365         # used when /analytics/allTime returns 403
# ====================================================

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# ======== APPLICATION TOKEN (Graph-only, non-interactive) ========
# Required admin-consented app permissions on Microsoft Graph:
#   Sites.Read.All, Files.Read.All
# (Optional) for group member count when site is group-connected: Group.Read.All
def get_graph_token() -> str:
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(result.get("error_description") or result)
    token = result["access_token"]
    if token.count(".") != 2:
        raise RuntimeError("Received token that doesn't look like a JWT.")
    return token

graph_token = get_graph_token()
H = {"Authorization": f"Bearer {graph_token}", "Accept": "application/json"}

# ---------- Helpers (Graph) ----------
def gget(url, headers=None):
    hdrs = dict(H)
    if headers:
        hdrs.update(headers)
    r = requests.get(url, headers=hdrs, proxies=PROXIES, verify=VERIFY)
    if r.status_code == 429:
        time.sleep(int(r.headers.get("Retry-After", "2")))
        r = requests.get(url, headers=hdrs, proxies=PROXIES, verify=VERIFY)
    r.raise_for_status()
    return r.json()

# 1) Resolve site
site = gget(f"https://graph.microsoft.com/v1.0/sites/{SITE_COMPOSITE_ID}?$select=id,webUrl,displayName")
site_id = site["id"]
print("Using site:", site["webUrl"])

# 2) Get the 'Site Pages' list GUID via Graph (no SharePoint _api)
lists = gget(
    f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists"
    f"?$filter=displayName eq '{SITE_PAGES_TITLE}'&$select=id,displayName"
)
vals = lists.get("value", [])
if not vals:
    raise RuntimeError(f"Could not find a list named '{SITE_PAGES_TITLE}' via Graph.")
list_guid = vals[0]["id"]

# 2a) List's driveId via Graph
drive = gget(f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_guid}/drive")
drive_id = drive["id"]
print("Drive ID:", drive_id)

# 2b) Group membership count (Graph only, robust)
group_members_count = 0
group_guid = None

# Try 1: drive.owner.group.id (works for most group-connected sites)
try:
    drv_meta = gget(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive?$select=owner")
    group_guid = ((drv_meta.get("owner") or {}).get("group") or {}).get("id")
except Exception:
    group_guid = None

# Try 2: sharepointIds.groupId (some tenants expose it here)
if not group_guid:
    try:
        ids = gget(f"https://graph.microsoft.com/v1.0/sites/{site_id}?$select=sharepointIds")
        group_guid = (ids.get("sharepointIds") or {}).get("groupId")
    except Exception:
        group_guid = None

# If we found a group, count members
if group_guid:
    try:
        members = gget(
            f"https://graph.microsoft.com/v1.0/groups/{group_guid}/members?$count=true&$top=1",
            headers={"ConsistencyLevel": "eventual"}
        )
        group_members_count = int(members.get("@odata.count", 0))
    except Exception:
        group_members_count = 0

print("Group GUID:", group_guid or "(none)")
print("Group members (Graph):", group_members_count)


# 3) Recursively enumerate all .aspx in Site Pages (Graph)
def children(drive_id, item_id="root"):
    url = (f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children?$top=200"
           if item_id == "root" else
           f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children?$top=200")
    items = []
    while url:
        j = gget(url)
        items += j.get("value", [])
        url = j.get("@odata.nextLink")
    return items

pages, stack, seen = [], ["root"], set()
while stack:
    node = stack.pop()
    for ch in children(drive_id, node):
        if ch["id"] in seen:
            continue
        seen.add(ch["id"])
        if "folder" in ch:
            stack.append(ch["id"])
        elif ch.get("name", "").lower().endswith(".aspx"):
            pages.append({"id": ch["id"], "name": ch["name"], "url": ch.get("webUrl")})
print(f"Found {len(pages)} .aspx page(s)")

# 4) Windows
def iso(dt): return dt.strftime("%Y-%m-%dT%H:%M:%SZ")
end_utc   = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(days=1)
start_daily = end_utc - timedelta(days=DAYS_DAILY)
start_fallback = end_utc - timedelta(days=FALLBACK_TOTAL_DAYS)

# 5) Collect analytics (daily + all-time with 403 fallback) — Graph endpoints
rows_daily, rows_total = [], []

for p in pages:
    # Daily (last N days)
    u = (f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{p['id']}"
         f"/getActivitiesByInterval(startDateTime='{iso(start_daily)}',endDateTime='{iso(end_utc)}',interval='day')")
    r = requests.get(u, headers=H, proxies=PROXIES, verify=VERIFY)
    if r.ok:
        for stat in r.json().get("value", []):
            acc = stat.get("access") or {}
            rows_daily.append({
                "Date": (stat.get("startDateTime") or "")[:10],
                "PageName": p["name"],
                "URL": p["url"],
                "Views": acc.get("actionCount", 0),
                "UniqueViewers": acc.get("actorCount", 0),
            })

    # All-time totals with 403-safe fallback
    try:
        at = gget(f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{p['id']}/analytics/allTime")
        acc = (at or {}).get("access") or {}
        rows_total.append({
            "PageName": p["name"], "URL": p["url"],
            "Views": acc.get("actionCount", 0),
            "UniqueViewers": acc.get("actorCount", 0),
            "AllTimeSource": "allTime"
        })
    except requests.HTTPError as e:
        if e.response.status_code == 403:
            u2 = (f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{p['id']}"
                  f"/getActivitiesByInterval(startDateTime='{iso(start_fallback)}',endDateTime='{iso(end_utc)}',interval='day')")
            r2 = requests.get(u2, headers=H, proxies=PROXIES, verify=VERIFY)
            views = uniq = 0
            if r2.ok:
                for stat in r2.json().get("value", []):
                    acc2 = stat.get("access") or {}
                    views += int(acc2.get("actionCount", 0) or 0)
                    uniq  += int(acc2.get("actorCount", 0) or 0)
            rows_total.append({
                "PageName": p["name"], "URL": p["url"],
                "Views": views,
                "UniqueViewers": uniq,
                "AllTimeSource": f"sum_{FALLBACK_TOTAL_DAYS}d"
            })
        else:
            raise

# Map the last segment of SITE_PATH to a short code
SITE_SEG  = SITE_PATH.rsplit('/', 1)[-1]  # e.g., "Project-SDEManagement"
SITE_CODE = {
    "ProjectManagement":       "PM",
    "Project-SDEManagement":   "SDE",
    "Project-SAAManagement":   "SAA",
    "QS":                      "QS",
}.get(SITE_SEG, "UNK")  # fallback if an unknown site

# Excel sheet names max 31 chars; keep it safe
def sheet(title: str) -> str:
    return (f"{title} {SITE_CODE}")[:31]

# 6) DataFrames & Excel
df_daily = (pd.DataFrame(rows_daily)
              .sort_values(["Date","Views"], ascending=[True, False])
              if rows_daily else pd.DataFrame(columns=["Date","PageName","URL","Views","UniqueViewers"]))

df_total = (pd.DataFrame(rows_total)
              .sort_values(["Views","UniqueViewers"], ascending=[False, False])
              if rows_total else pd.DataFrame(columns=["PageName","URL","Views","UniqueViewers","AllTimeSource"]))

# Add group membership count to every row in AllTimeTotals
if not df_total.empty:
    df_total["GroupMembers"] = group_members_count
else:
    df_total["GroupMembers"] = pd.Series(dtype="int64")

xlsx = f"site_pages_analytics_{SITE_CODE}.xlsx"  # e.g., site_pages_analytics_SDE.xlsx

with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
    df_daily.to_excel(xw, sheet_name=sheet("DailyLast90"), index=False)
    df_total.to_excel(xw, sheet_name=sheet("AllTimeTotals"), index=False)

print("Saved →", xlsx, "| Sheets:", sheet("DailyLast90"), ",", sheet("AllTimeTotals"))
