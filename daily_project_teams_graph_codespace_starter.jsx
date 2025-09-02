// README (quick-start is below; full instructions in comments throughout the files)
// ─────────────────────────────────────────────────────────────────────────────
// This Codespace scaffolds a FastAPI + MSAL service that exposes simple REST
// endpoints to call Microsoft Graph for Teams (channel/chat messages), Outlook
// Mail.Send, and Calendar events using **application (client‑credentials)** auth.
//
// IMPORTANT SECURITY NOTES
// • Do NOT commit secrets. Use GitHub Codespaces “Secrets” or a local .env
//   that you never commit. Since credentials were pasted in chat, ROTATE them.
// • Principle of least privilege: only grant the Graph permissions you need.
// • Prefer private repo while bootstrapping, then migrate to Enterprise.
//
// QUICK START (once app registration is ready; see entra-setup.md below):
// 1) Add Codespace Secrets: TENANT_ID, CLIENT_ID, CLIENT_SECRET.
// 2) (Optional) Also add TEAM_ID, CHANNEL_ID, SENDER_UPN for quicker tests.
// 3) Open this repo in Codespaces → it builds the Dev Container.
// 4) App runs on port 8080. Test endpoints with curl or VS Code REST.
// 5) Try:
//    curl -X POST "${CODESPACE_NAME}-8080.app.github.dev/send-channel-message" \
//      -H 'Content-Type: application/json' \
//      -d '{"team_id":"<TEAM_ID>","channel_id":"<CHANNEL_ID>","message":"Hello from Codespaces!"}'
// ─────────────────────────────────────────────────────────────────────────────

// .devcontainer/devcontainer.json
// Place this file at: .devcontainer/devcontainer.json
{
  "name": "teams-graph-codespace",
  "features": {
    "ghcr.io/devcontainers/features/python:1": {
      "version": "3.11"
    }
  },
  "forwardPorts": [8080],
  "portsAttributes": {
    "8080": { "label": "FastAPI" }
  },
  "containerEnv": {
    "PYTHONDONTWRITEBYTECODE": "1",
    "PYTHONUNBUFFERED": "1"
  },
  "postCreateCommand": "pip install -r requirements.txt",
  "remoteUser": "vscode"
}

// requirements.txt
fastapi==0.112.2
uvicorn==0.30.6
msal==1.28.0
httpx==0.27.2
python-dotenv==1.0.1
pydantic==2.9.2

// Dockerfile (optional for local docker run; Codespace uses devcontainer)
// docker build -t teams-graph:latest . && docker run -p 8080:8080 --env-file .env teams-graph
FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt
COPY app ./app
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8080"]

// .env.sample  (NEVER COMMIT a real .env)
TENANT_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
CLIENT_ID="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
CLIENT_SECRET="<client-secret>"
// Optional convenience defaults for quick tests:
TEAM_ID="<team-id>"
CHANNEL_ID="<channel-id>"
SENDER_UPN="someone@yourtenant.onmicrosoft.com"

// app/main.py
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, EmailStr
import os
import httpx
import msal
from dotenv import load_dotenv

load_dotenv()

GRAPH_SCOPE = ["https://graph.microsoft.com/.default"]
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
DEFAULT_TEAM_ID = os.getenv("TEAM_ID")
DEFAULT_CHANNEL_ID = os.getenv("CHANNEL_ID")
DEFAULT_SENDER_UPN = os.getenv("SENDER_UPN")

if not (TENANT_ID and CLIENT_ID and CLIENT_SECRET):
    raise RuntimeError("TENANT_ID, CLIENT_ID, CLIENT_SECRET must be set via env or Codespace Secrets")

app = FastAPI(title="Teams/Graph API Bridge", version="0.1.0")

# MSAL app (client-credentials)
confidential_app = msal.ConfidentialClientApplication(
    CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    client_credential=CLIENT_SECRET,
)

def _get_token() -> str:
    result = confidential_app.acquire_token_silent(GRAPH_SCOPE, account=None)
    if not result:
        result = confidential_app.acquire_token_for_client(scopes=GRAPH_SCOPE)
    if "access_token" not in result:
        raise HTTPException(status_code=500, detail=f"MSAL auth error: {result.get('error_description')}")
    return result["access_token"]

@app.get("/health")
def health():
    return {"ok": True}

# MODELS
class ChannelMessageReq(BaseModel):
    team_id: str | None = None
    channel_id: str | None = None
    message: str

class EmailReq(BaseModel):
    sender_upn: str | None = None  # which mailbox to send from (app perms)
    to: list[EmailStr]
    subject: str
    body_html: str

class MeetingReq(BaseModel):
    organizer_upn: str | None = None
    subject: str
    start_iso: str  # e.g., "2025-09-02T15:00:00"
    end_iso: str    # e.g., "2025-09-02T15:30:00"
    attendees: list[EmailStr] = []
    body_html: str | None = None
    location: str | None = None

# HELPERS
async def graph_post(url: str, json: dict):
    token = _get_token()
    async with httpx.AsyncClient(timeout=30) as client:
        r = await client.post(url, headers={"Authorization": f"Bearer {token}"}, json=json)
        if r.status_code >= 400:
            raise HTTPException(status_code=r.status_code, detail=r.text)
        if r.text:
            return r.json()
        return {"ok": True}

async def graph_get(url: str):
    token = _get_token()
    async with httpx.AsyncClient(timeout=30) as client:
        r = await client.get(url, headers={"Authorization": f"Bearer {token}"})
        if r.status_code >= 400:
            raise HTTPException(status_code=r.status_code, detail=r.text)
        return r.json()

# ENDPOINTS
@app.post("/send-channel-message")
async def send_channel_message(req: ChannelMessageReq):
    team_id = req.team_id or DEFAULT_TEAM_ID
    channel_id = req.channel_id or DEFAULT_CHANNEL_ID
    if not team_id or not channel_id:
        raise HTTPException(status_code=400, detail="team_id and channel_id are required (env or body)")
    url = f"{GRAPH_BASE}/teams/{team_id}/channels/{channel_id}/messages"
    payload = {"body": {"content": req.message}}
    return await graph_post(url, payload)

@app.post("/send-email")
async def send_email(req: EmailReq):
    sender = req.sender_upn or DEFAULT_SENDER_UPN
    if not sender:
        raise HTTPException(status_code=400, detail="sender_upn required (env or body)")
    url = f"{GRAPH_BASE}/users/{sender}/sendMail"
    message = {
        "message": {
            "subject": req.subject,
            "body": {"contentType": "HTML", "content": req.body_html},
            "toRecipients": [{"emailAddress": {"address": x}} for x in req.to],
        },
        "saveToSentItems": True,
    }
    return await graph_post(url, message)

@app.post("/schedule-meeting")
async def schedule_meeting(req: MeetingReq):
    organizer = req.organizer_upn or DEFAULT_SENDER_UPN
    if not organizer:
        raise HTTPException(status_code=400, detail="organizer_upn required (env or body)")
    url = f"{GRAPH_BASE}/users/{organizer}/events"
    attendees = [{
        "emailAddress": {"address": a},
        "type": "required"
    } for a in req.attendees]
    body = {
        "subject": req.subject,
        "start": {"dateTime": req.start_iso, "timeZone": "UTC"},
        "end":   {"dateTime": req.end_iso,   "timeZone": "UTC"},
        "location": {"displayName": req.location} if req.location else None,
        "body": {"contentType": "HTML", "content": req.body_html or ""},
        "attendees": attendees,
        "isOnlineMeeting": True,
        "onlineMeetingProvider": "teamsForBusiness"
    }
    # Remove Nones for cleanliness
    body = {k:v for k,v in body.items() if v is not None}
    return await graph_post(url, body)

# Utility endpoints to discover IDs
@app.get("/teams")
async def list_teams():
    # List all M365 Groups that are Teams, then expand team metadata
    groups = await graph_get(f"{GRAPH_BASE}/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')")
    return groups

@app.get("/teams/{team_id}/channels")
async def list_channels(team_id: str):
    return await graph_get(f"{GRAPH_BASE}/teams/{team_id}/channels")

// entra-setup.md (concise)
# Entra / Graph app registration (client-credentials)
1. **Register app**: Entra ID → App registrations → *New registration* → Name `daily-project-graph-daemon` → *Accounts in this org directory only* → Register.
2. **Add a client secret**: Certificates & secrets → *New client secret* → copy value.
3. **API permissions (Application)** → *Add a permission* → Microsoft Graph → *Application permissions* → add:
   - `ChannelMessage.Send` *(post to Teams channels)*
   - `Chat.ReadWrite.All` *(send to 1:1 / group chats if needed)*
   - `Mail.Send` *(send email)*
   - `Calendars.ReadWrite` *(create meetings)*
   - `User.Read.All` *(resolve users, optional but useful)*
   - `Group.Read.All` *(enumerate groups/teams; Group.ReadWrite.All if creating channels)*
   Click **Grant admin consent** for the tenant.
4. **Collect IDs**: *Overview* → copy **Application (client) ID** and **Directory (tenant) ID**.
5. **Create/choose a Team & Channel** for testing. Capture their IDs:
   - Use this service: `GET /groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')` → pick a group `id`.
   - Get channels: `GET /teams/{team-id}/channels` → pick a channel `id` (e.g., `general`).
6. **Codespaces Secrets**: In GitHub → Codespaces → *Secrets* → add `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`. Optionally `TEAM_ID`, `CHANNEL_ID`, `SENDER_UPN`.

# Test commands
curl -X GET "$BASE/health"

curl -X POST "$BASE/send-channel-message" -H 'Content-Type: application/json' \
  -d '{"team_id":"<TEAM_ID>","channel_id":"<CHANNEL_ID>","message":"Hello from Codespaces"}'

curl -X POST "$BASE/send-email" -H 'Content-Type: application/json' \
  -d '{"sender_upn":"<user@tenant.onmicrosoft.com>","to":["<recipient@tenant>"] ,"subject":"Test","body_html":"<b>Hi</b> from Codespaces"}'

curl -X POST "$BASE/schedule-meeting" -H 'Content-Type: application/json' \
  -d '{"organizer_upn":"<user@tenant>","subject":"Daily Standup","start_iso":"2025-09-02T15:00:00","end_iso":"2025-09-02T15:15:00","attendees":["<recipient@tenant>"],"body_html":"Auto-sent.","location":"Teams"}'

// ─────────────────────────────────────────────────────────────────────────────
// Delegated (Device Code) add‑on — paste below into app/main.py
// ─────────────────────────────────────────────────────────────────────────────
// Enables posting to Teams as the signed‑in user, and optional mail/calendar.
// Entra → App → Authentication → **Allow public client flows = Yes**.
// Codespaces Secrets needed: TENANT_ID, CLIENT_ID (CLIENT_SECRET not used here).

# --- Delegated auth (Device Code) ---
import json
from msal import PublicClientApplication, SerializableTokenCache

DELEGATED_SCOPES = [
    "https://graph.microsoft.com/User.Read",
    "https://graph.microsoft.com/ChannelMessage.Send",
    "https://graph.microsoft.com/Chat.ReadWrite",
    "https://graph.microsoft.com/Mail.Send",
    "https://graph.microsoft.com/Calendars.ReadWrite",
    "https://graph.microsoft.com/Group.Read.All",
    "https://graph.microsoft.com/Group.ReadWrite.All",
    "offline_access",
]

CACHE_PATH = "/workspaces/token-cache.json"

token_cache = SerializableTokenCache()
try:
    with open(CACHE_PATH, "r") as f:
        token_cache.deserialize(f.read())
except FileNotFoundError:
    pass


def save_cache():
    if token_cache.has_state_changed:
        with open(CACHE_PATH, "w") as f:
            f.write(token_cache.serialize())


public_app = PublicClientApplication(
    CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    token_cache=token_cache,
)


def get_user_token() -> str:
    accounts = public_app.get_accounts()
    if accounts:
        result = public_app.acquire_token_silent(DELEGATED_SCOPES, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]
    raise HTTPException(status_code=401, detail="No delegated token—run /auth/device-login")


@app.post("/auth/device-login")
def device_login():
    flow = public_app.initiate_device_flow(scopes=DELEGATED_SCOPES)
    if "user_code" not in flow:
        raise HTTPException(status_code=500, detail="Failed to create device code flow")
    # NOTE: MSAL will print the login message to logs; we return basics too
    result = public_app.acquire_token_by_device_flow(flow)
    save_cache()
    if "access_token" not in result:
        raise HTTPException(status_code=401, detail=f"Device code auth failed: {result.get('error_description')}")
    acct = public_app.get_accounts()[0].get("username")
    return {"ok": True, "account": acct}


@app.post("/logout")
def logout():
    for acc in public_app.get_accounts():
        public_app.remove_account(acc)
    save_cache()
    return {"ok": True}


async def graph_get_me_delegated(token: str):
    async with httpx.AsyncClient(timeout=30) as client:
        r = await client.get(f"{GRAPH_BASE}/me", headers={"Authorization": f"Bearer {token}"})
        if r.status_code >= 400:
            raise HTTPException(status_code=r.status_code, detail=r.text)
        return r.json()


@app.get("/me-delegated")
async def me_delegated():
    token = get_user_token()
    return await graph_get_me_delegated(token)


@app.post("/send-channel-message-delegated")
async def send_channel_message_delegated(req: ChannelMessageReq):
    team_id = req.team_id or DEFAULT_TEAM_ID
    channel_id = req.channel_id or DEFAULT_CHANNEL_ID
    if not team_id or not channel_id:
        raise HTTPException(status_code=400, detail="team_id and channel_id are required")
    token = get_user_token()
    url = f"{GRAPH_BASE}/teams/{team_id}/channels/{channel_id}/messages"
    payload = {"body": {"content": req.message}}
    async with httpx.AsyncClient(timeout=30) as client:
        r = await client.post(url, headers={"Authorization": f"Bearer {token}"}, json=payload)
        if r.status_code >= 400:
            raise HTTPException(status_code=r.status_code, detail=r.text)
        return r.json() if r.text else {"ok": True}


@app.post("/send-email-delegated")
async def send_email_delegated(req: EmailReq):
    token = get_user_token()
    me = await graph_get_me_delegated(token)
    sender = me["userPrincipalName"]
    url = f"{GRAPH_BASE}/users/{sender}/sendMail"
    message = {
        "message": {
            "subject": req.subject,
            "body": {"contentType": "HTML", "content": req.body_html},
            "toRecipients": [{"emailAddress": {"address": x}} for x in req.to],
        },
        "saveToSentItems": True,
    }
    async with httpx.AsyncClient(timeout=30) as client:
        r = await client.post(url, headers={"Authorization": f"Bearer {token}"}, json=message)
        if r.status_code >= 400:
            raise HTTPException(status_code=r.status_code, detail=r.text)
        return r.json() if r.text else {"ok": True}


@app.post("/schedule-meeting-delegated")
async def schedule_meeting_delegated(req: MeetingReq):
    token = get_user_token()
    me = await graph_get_me_delegated(token)
    organizer_upn = req.organizer_upn or me["userPrincipalName"]
    url = f"{GRAPH_BASE}/users/{organizer_upn}/events"
    attendees = [{"emailAddress": {"address": a}, "type": "required"} for a in req.attendees]
    body = {
        "subject": req.subject,
        "start": {"dateTime": req.start_iso, "timeZone": "UTC"},
        "end":   {"dateTime": req.end_iso,   "timeZone": "UTC"},
        "location": {"displayName": req.location} if req.location else None,
        "body": {"contentType": "HTML", "content": req.body_html or ""},
        "attendees": attendees,
        "isOnlineMeeting": True,
        "onlineMeetingProvider": "teamsForBusiness",
    }
    body = {k: v for k, v in body.items() if v is not None}
    async with httpx.AsyncClient(timeout=30) as client:
        r = await client.post(url, headers={"Authorization": f"Bearer {token}"}, json=body)
        if r.status_code >= 400:
            raise HTTPException(status_code=r.status_code, detail=r.text)
        return r.json() if r.text else {"ok": True}


// END OF STARTER
