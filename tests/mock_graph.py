"""Mock Microsoft Entra ID + Graph (beta) + Azure Blob endpoints for the Intune uploader test.

Usage: python3 tests/mock_graph.py <port> <state-dir>
"""
import json, os, sys, threading, base64, re
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs

PORT = int(sys.argv[1]); STATE_DIR = sys.argv[2]
os.makedirs(STATE_DIR, exist_ok=True)
lock = threading.Lock()
S = {"apps": {}, "content_versions": 0, "files": {}, "blocks": {}, "commits": [], "patches": [], "device_polls": 0, "noperm_tokens": 0}

def jwt(claims):
    enc = lambda d: base64.urlsafe_b64encode(json.dumps(d).encode()).decode().rstrip("=")
    return enc({"alg": "none", "typ": "JWT"}) + "." + enc(claims) + ".sig"

# An app registration without Intune permission: Entra ID issues a token without roles,
# Intune answers like the real service does (401, ErrorCode Forbidden, Message as JSON string).
NOPERM_TOKEN_PREFIX = "Bearer " + jwt({"tid": "tenant-1", "appid": "client-noperm", "roles": []})[:20]
INTUNE_FORBIDDEN = {"ErrorCode": "Forbidden", "Message": json.dumps({"_version": 3, "Message": "An error has occurred - Operation ID (for customer support): 00000000-0000-0000-0000-000000000000 - Activity ID: 96cb80e2-75df-46eb-9b03-3ee8517a9c75 - Url: https://proxy.msub09.manage.microsoft.com/AppLifecycle_2611/StatelessAppMetadataFEService/deviceAppManagement/mobileApps?api-version=5026-08-19", "CustomApiErrorPhrase": "", "RetryAfter": None, "ErrorSourceService": "", "HttpHeaders": "{\"WWW-Authenticate\":\"Bearer\"}"}, indent=2), "Target": None, "Details": None, "InnerError": None, "InstanceAnnotations": []}

class H(BaseHTTPRequestHandler):
    def log_message(self, *a): pass
    def body(self):
        n = int(self.headers.get("Content-Length") or 0)
        return self.rfile.read(n) if n else b""
    def send(self, code, obj=None, raw=None):
        data = raw if raw is not None else (json.dumps(obj).encode() if obj is not None else b"")
        self.send_response(code)
        if data: self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers(); self.wfile.write(data)
    def authed(self):
        if (self.headers.get("Authorization") or "").startswith(NOPERM_TOKEN_PREFIX):
            self.send(401, INTUNE_FORBIDDEN); return False
        if self.headers.get("Authorization") != "Bearer tok-1":
            self.send(401, {"error": {"code": "InvalidAuthenticationToken", "message": "Access token is empty."}}); return False
        return True
    def do_POST(self):
        u = urlparse(self.path); p = u.path; b = self.body()
        with lock:
            if p.endswith("/oauth2/v2.0/token"):
                f = {k: v[0] for k, v in parse_qs(b.decode()).items()}
                if f.get("grant_type") == "client_credentials":
                    if f.get("client_secret") == "secret-1" and f.get("scope") == "https://graph.microsoft.com/.default":
                        return self.send(200, {"access_token": "tok-1", "expires_in": 3599, "token_type": "Bearer"})
                    if f.get("client_secret") == "secret-noperm":
                        S["noperm_tokens"] += 1
                        return self.send(200, {"access_token": jwt({"tid": "tenant-1", "appid": "client-noperm", "roles": []}), "expires_in": 3599, "token_type": "Bearer"})
                    return self.send(401, {"error": "invalid_client", "error_description": "AADSTS7000215: Invalid client secret provided.\r\nTrace ID: x"})
                if f.get("grant_type", "").endswith("device_code"):
                    S["device_polls"] += 1
                    if S["device_polls"] < 2:
                        return self.send(400, {"error": "authorization_pending", "error_description": "pending"})
                    return self.send(200, {"access_token": "tok-1", "refresh_token": "rt-1", "expires_in": 3599})
                return self.send(400, {"error": "unsupported_grant_type"})
            if p.endswith("/oauth2/v2.0/devicecode"):
                return self.send(200, {"device_code": "dc-1", "user_code": "ABCD-EFGH", "verification_uri": "https://microsoft.com/devicelogin", "expires_in": 900, "interval": 1})
            if not self.authed(): return
            j = json.loads(b or b"{}")
            if p == "/beta/deviceAppManagement/mobileApps":
                for k in ["displayName", "description", "publisher", "installCommandLine", "uninstallCommandLine", "setupFilePath", "fileName", "rules", "installExperience"]:
                    if not j.get(k): return self.send(400, {"error": {"code": "BadRequest", "message": f"missing {k}"}})
                app_id = f"app-{len(S['apps']) + 1}"; j["id"] = app_id; S["apps"][app_id] = j
                return self.send(201, j)
            m = re.match(r"^/beta/deviceAppManagement/mobileApps/([^/]+)/microsoft.graph.win32LobApp/contentVersions$", p)
            if m:
                S["content_versions"] += 1
                return self.send(201, {"id": str(S["content_versions"])})
            m = re.match(r"^/beta/deviceAppManagement/mobileApps/([^/]+)/microsoft.graph.win32LobApp/contentVersions/([^/]+)/files$", p)
            if m:
                if j.get("@odata.type") != "#microsoft.graph.mobileAppContentFile" or not j.get("sizeEncrypted"):
                    return self.send(400, {"error": {"message": "bad file body"}})
                fid = f"file-{len(S['files']) + 1}"
                S["files"][fid] = {"state": "azureStorageUriRequestPending", "gets": 0, "body": j, "app": m.group(1), "cv": m.group(2)}
                return self.send(201, {"id": fid, "uploadState": "azureStorageUriRequestPending"})
            m = re.match(r"^/beta/.*/files/([^/]+)/commit$", p)
            if m:
                fid = m.group(1); info = j.get("fileEncryptionInfo", {})
                S["commits"].append({"file": fid, "info": info})
                S["files"][fid]["state"] = "commitFilePending"; S["files"][fid]["gets"] = 0
                return self.send(200)
            return self.send(404, {"error": {"message": "no route " + p}})
    def do_PATCH(self):
        p = urlparse(self.path).path; b = self.body()
        with lock:
            if not self.authed(): return
            m = re.match(r"^/beta/deviceAppManagement/mobileApps/([^/]+)$", p)
            if m and m.group(1) in S["apps"]:
                j = json.loads(b); S["patches"].append({"app": m.group(1), "body": j}); S["apps"][m.group(1)].update(j)
                return self.send(204)
            return self.send(404, {"error": {"message": "no app"}})
    def do_GET(self):
        u = urlparse(self.path); p = u.path
        with lock:
            if p == "/__state":
                return self.send(200, S)
            if not self.authed(): return
            if p == "/beta/deviceAppManagement/mobileApps":
                if "isof(" not in parse_qs(u.query).get("$filter", [""])[0]:
                    return self.send(400, {"error": {"message": "filter expected"}})
                return self.send(200, {"value": [{"id": k, "displayName": v["displayName"], "displayVersion": v.get("displayVersion", ""), "publisher": v["publisher"]} for k, v in S["apps"].items()]})
            m = re.match(r"^/beta/.*/files/([^/]+)$", p)
            if m:
                f = S["files"][m.group(1)]; f["gets"] += 1
                if f["state"] == "azureStorageUriRequestPending" and f["gets"] >= 2:
                    f["state"] = "azureStorageUriRequestSuccess"
                elif f["state"] == "commitFilePending" and f["gets"] >= 2:
                    f["state"] = "commitFileSuccess"
                out = {"id": m.group(1), "uploadState": f["state"]}
                if f["state"] == "azureStorageUriRequestSuccess":
                    out["azureStorageUri"] = f"http://127.0.0.1:{PORT}/blob/{m.group(1)}?sv=2024&sig=abc%2Bdef"
                return self.send(200, out)
            return self.send(404, {"error": {"message": "no route " + p}})
    def do_PUT(self):
        u = urlparse(self.path); q = parse_qs(u.query); b = self.body()
        m = re.match(r"^/blob/([^/]+)$", u.path)
        with lock:
            if not m or q.get("sig") != ["abc+def"]: return self.send(403, raw=b"bad sas")
            fid = m.group(1)
            if q.get("comp") == ["block"]:
                if self.headers.get("x-ms-blob-type") != "BlockBlob": return self.send(400, raw=b"blob type")
                S["blocks"].setdefault(fid, {})[q["blockid"][0]] = len(b)
                with open(os.path.join(STATE_DIR, f"{fid}.{q['blockid'][0]}"), "wb") as fh: fh.write(b)
                return self.send(201, raw=b"")
            if q.get("comp") == ["blocklist"]:
                ids = re.findall(r"<Latest>(.*?)</Latest>", b.decode())
                with open(os.path.join(STATE_DIR, f"{fid}.blob"), "wb") as out:
                    for i in ids:
                        with open(os.path.join(STATE_DIR, f"{fid}.{i}"), "rb") as fh: out.write(fh.read())
                S["files"][fid]["blocklist"] = ids
                return self.send(201, raw=b"")
            return self.send(400, raw=b"comp?")

ThreadingHTTPServer(("127.0.0.1", PORT), H).serve_forever()
