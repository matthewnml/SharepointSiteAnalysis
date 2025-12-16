import msal, requests, pandas as pd, time
from datetime import datetime, timedelta, timezone



# ==== inputs (change these when switching sites) ====
TENANT_ID = ""
CLIENT_ID = ""
HOSTNAME  = ""

# SITE PATHS & IDS
SITE_PATH = ""  # change only this line when switching sites
SITE_COMPOSITE_ID = ""
# If you switch to SDE later, only change the two lines above to:
# SITE_PATH = "sites/Project-SDEManagement"
# SITE_COMPOSITE_ID = "tiongseng.sharepoint.com,b6c68d09-2c2c-4a8c-8557-1217bd3506bb,1fe88523-7a1b-4cfb-be57-00d05b591233"

SITE_PAGES_TITLE = "Site Pages"   # change only if the library is named differently
DAYS_DAILY = 90                   # daily window
FALLBACK_TOTAL_DAYS = 365         # used when /analytics/allTime returns 403
# ====================================================

authority = f"https://login.microsoftonline.com/{TENANT_ID}"

# Graph token (delegated)
graph_scopes = ["Files.Read.All", "Sites.Read.All", "User.Read"]
app = msal.PublicClientApplication(CLIENT_ID, authority=authority)
flow = app.initiate_device_flow(scopes=graph_scopes); print(flow["message"])
res = app.acquire_token_by_device_flow(flow); assert "access_token" in res, res
H = {"Authorization": f"Bearer {res['access_token']}"}

# SharePoint token (REST) to fetch the Site Pages list GUID
sp_scopes = [f"https://{HOSTNAME}/AllSites.Read"]
flow2 = app.initiate_device_flow(scopes=sp_scopes); print(flow2["message"])
res2 = app.acquire_token_by_device_flow(flow2); assert "access_token" in res2, res2
S = {"Authorization": f"Bearer {res2['access_token']}", "Accept":"application/json;odata=nometadata"}

def gget(url):
    r = requests.get(url, headers=H)
    if r.status_code == 429:
        time.sleep(int(r.headers.get("Retry-After", "2")))
        r = requests.get(url, headers=H)
    r.raise_for_status()
    return r.json()

def sget(url):
    r = requests.get(url, headers=S); r.raise_for_status(); return r.json()

# 1) Resolve site
site = gget(f"https://graph.microsoft.com/v1.0/sites/{SITE_COMPOSITE_ID}?$select=id,webUrl")
site_id = site["id"]; print("Using site:", site["webUrl"])

# 2) Site Pages → list GUID (REST) → driveId (Graph)
lst = sget(f"https://{HOSTNAME}/{SITE_PATH}/_api/web/lists/getByTitle('{SITE_PAGES_TITLE}')"
           "?$select=Id,Title,RootFolder/ServerRelativeUrl&$expand=RootFolder")
list_guid = lst["Id"]
drive = gget(f"https://graph.microsoft.com/v1.0/sites/{site_id}/lists/{list_guid}/drive")
drive_id = drive["id"]; print("Drive ID:", drive_id)

# 2b) Get group membership count (M365 group if present, else SP 'Site Members')
def gget_with_headers(url, extra_headers=None):
    hdrs = dict(H)
    if extra_headers:
        hdrs.update(extra_headers)
    r = requests.get(url, headers=hdrs)
    if r.status_code == 429:
        time.sleep(int(r.headers.get("Retry-After", "2")))
        r = requests.get(url, headers=hdrs)
    r.raise_for_status()
    return r.json()

# Ask SharePoint for the site's GroupId (non-zero only if group-connected)
site_meta = sget(f"https://{HOSTNAME}/{SITE_PATH}/_api/site?$select=GroupId")
group_guid = (site_meta or {}).get("GroupId", "") or ""
group_members_count = None

if group_guid and group_guid != "00000000-0000-0000-0000-000000000000":
    # Use Graph $count for members (requires ConsistencyLevel:eventual)
    try:
        j = gget_with_headers(
            f"https://graph.microsoft.com/v1.0/groups/{group_guid}/members?$count=true&$top=1",
            {"ConsistencyLevel": "eventual"}
        )
        group_members_count = int(j.get("@odata.count", 0))
    except Exception:
        group_members_count = None

# Fallback: count users in the SharePoint "Site Members" group
if group_members_count is None:
    try:
        m = sget(f"https://{HOSTNAME}/{SITE_PATH}/_api/web/AssociatedMemberGroup/Users?$top=5000")
        group_members_count = len(m.get("value", []))
    except Exception:
        group_members_count = 0  # last-resort default

print("Group members:", group_members_count)

# 3) Recursively enumerate all .aspx in Site Pages
def children(drive_id, item_id="root"):
    url = (f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children?$top=200"
           if item_id == "root" else
           f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/children?$top=200")
    items = []
    while url:
        j = gget(url); items += j.get("value", []); url = j.get("@odata.nextLink")
    return items

pages, stack, seen = [], ["root"], set()
while stack:
    node = stack.pop()
    for ch in children(drive_id, node):
        if ch["id"] in seen: continue
        seen.add(ch["id"])
        if "folder" in ch:
            stack.append(ch["id"])
        elif ch.get("name","").lower().endswith(".aspx"):
            pages.append({"id": ch["id"], "name": ch["name"], "url": ch.get("webUrl")})
print(f"Found {len(pages)} .aspx page(s)")

# 4) Windows
def iso(dt): return dt.strftime("%Y-%m-%dT%H:%M:%SZ")
end_utc   = datetime.now(timezone.utc).replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(days=1)
start_daily = end_utc - timedelta(days=DAYS_DAILY)
start_fallback = end_utc - timedelta(days=FALLBACK_TOTAL_DAYS)

# 5) Collect analytics (daily + all-time with 403 fallback)
rows_daily, rows_total = [], []

for p in pages:
    # Daily (last N days)
    u = (f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{p['id']}"
         f"/getActivitiesByInterval(startDateTime='{iso(start_daily)}',endDateTime='{iso(end_utc)}',interval='day')")
    r = requests.get(u, headers=H)
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
            r2 = requests.get(u2, headers=H)
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

