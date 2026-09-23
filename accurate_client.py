"""
Minimal Accurate Online client: log in, open the company database, run a saved report
and export it to .xlsx - the same requests the browser makes (see the HARs in
Data\\Sent Email\\1st_account.accurate.id.har / 2nd_iris.accurate.id.har).

Login (account.accurate.id):
    POST /pre-login.do   account, password="up"+b64({v,p,d})        -> d.permit
    POST /auth.do        j_username, j_password="ua"+b64({v,p,c,t,d}) -> 302 /manage
  (the encoding is nucauth.pre/post from cdn.accurate.id/.../js/hello/nucauth.js)

Database:
    POST /manage/database-list.do                                  -> d.dbList[]
    GET  {host}/accurate/open.do?uid={uniqueId}&product=aol        -> session, _dsi/_usi

Report ({report host}/accurate/...):
    POST report/init-report-input.do   id, planId, _dsi             -> d.reportInput
    POST report/bg-execute-report.do   id, planId, reportInput, ... -> b (background pid)
    POST company/bg-proc-response.do   bgPid, keepCache             -> status, cacheId
    POST report/export-report.do       cacheId, exportType=xls      -> .xlsx bytes
"""

from __future__ import annotations

import base64
import json
import re
import time
from urllib.parse import urlsplit

import requests

ACCOUNT = "https://account.accurate.id"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36")


class AccurateError(RuntimeError):
    pass


def _b64json(obj) -> str:
    return base64.b64encode(json.dumps(obj, separators=(",", ":")).encode()).decode()


class AccurateClient:
    def __init__(self, email: str, password: str, device_id: str, timeout: int = 60):
        self._email = email
        self._password = password
        self._device = device_id
        self.timeout = timeout
        self.s = requests.Session()
        self.s.headers.update({"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"})
        self.host = None        # e.g. https://iris.accurate.id
        self.report_host = None  # e.g. https://iris-report.accurate.id
        self.dsi = None
        self.usi = None

    # -- helpers ---------------------------------------------------------- #

    def _post_json(self, url: str, data: dict, referer: str | None = None) -> dict:
        headers = {"X-Requested-With": "XMLHttpRequest"}
        if referer:
            headers["Referer"] = referer
            headers["Origin"] = "{0.scheme}://{0.netloc}".format(urlsplit(referer))
        r = self.s.post(url, data=data, headers=headers, timeout=self.timeout)
        r.raise_for_status()
        try:
            j = r.json()
        except ValueError as exc:
            raise AccurateError(f"{url}: expected JSON, got {r.headers.get('content-type')}") from exc
        if not j.get("s", False):
            # "d" carries the server's error message, never a credential.
            raise AccurateError(f"{url.rsplit('/', 1)[-1]} failed: {j.get('d')}")
        return j

    # -- login ------------------------------------------------------------ #

    def login(self) -> None:
        self.s.get(ACCOUNT + "/", timeout=self.timeout)
        pre = self._post_json(ACCOUNT + "/pre-login.do", {
            "account": self._email,
            "password": "up" + _b64json({"v": 1, "p": self._password, "d": self._device}),
        }, referer=ACCOUNT + "/")["d"]
        if pre.get("requireTotp"):
            raise AccurateError("Accurate asks for a 2FA/TOTP code for this login - not supported.")
        if not pre.get("valid"):
            raise AccurateError(f"Login rejected: {pre.get('errorMsg') or pre.get('errorCode')}")
        j_password = "ua" + _b64json({"v": 1, "p": self._password, "c": pre["permit"],
                                      "t": None, "d": self._device})
        r = self.s.post(ACCOUNT + "/auth.do", data={"j_username": self._email, "j_password": j_password},
                        headers={"Referer": ACCOUNT + "/", "Origin": ACCOUNT},
                        timeout=self.timeout, allow_redirects=True)
        r.raise_for_status()
        if "/manage" not in r.url:
            raise AccurateError(f"Login did not reach /manage (ended at {urlsplit(r.url).path})")

    # -- database --------------------------------------------------------- #

    def databases(self) -> list[dict]:
        return self._post_json(ACCOUNT + "/manage/database-list.do", {},
                               referer=ACCOUNT + "/manage")["d"]["dbList"]

    def open_database(self, name_contains: str) -> dict:
        dbs = self.databases()
        match = [d for d in dbs if name_contains.lower() in str(d.get("name", "")).lower()]
        if len(match) != 1:
            names = [d.get("name") for d in dbs]
            raise AccurateError(f"Expected one database matching {name_contains!r}, found {len(match)}: {names}")
        db = match[0]
        if db.get("host"):
            url = (f"{db['host']}/accurate/open.do?id={db['id']}" if db.get("isPrivateCloud")
                   else f"{db['host']}/accurate/open.do?uid={db['uniqueId']}")
        else:
            url = f"{ACCOUNT}/manage/open.do?id={db['id']}"
        r = self.s.get(url + "&product=aol", headers={"Referer": ACCOUNT + "/manage"},
                       timeout=self.timeout, allow_redirects=True)
        r.raise_for_status()
        final = urlsplit(r.url)
        self.host = f"{final.scheme}://{final.netloc}"
        self._find_session_ids(r)
        return db

    def _find_session_ids(self, r: requests.Response) -> None:
        """The dashboard carries the _dsi/_usi session ids that every later call
        posts. Look for them in the final URL, the page, and the cookies."""
        text = r.text
        found = {}
        for name in ("_dsi", "_usi"):
            pats = [rf"[?&]{name}=([^&#\s\"']+)",
                    rf"['\"]?{name}['\"]?\s*[:=]\s*['\"]([^'\"]+)['\"]"]
            for src in (r.url, text):
                for p in pats:
                    m = re.search(p, src)
                    if m:
                        found[name] = m.group(1)
                        break
                if name in found:
                    break
            if name not in found:
                for c in self.s.cookies:
                    if c.name.lower().lstrip("_") == name.lstrip("_"):
                        found[name] = c.value
        self.dsi, self.usi = found.get("_dsi"), found.get("_usi")
        # The report calls go to a separate "-report" host next to the UI host.
        m = re.search(r"https://[\w.-]*-report\.accurate\.id", text)
        self.report_host = m.group(0) if m else self.host.replace(".accurate.id", "-report.accurate.id")
        if not (self.dsi and self.usi):
            raise AccurateError(
                "Opened the database but could not find the _dsi/_usi session ids "
                f"(landed on {final_path(r)}, {len(text)} bytes, cookies: "
                f"{sorted(c.name for c in self.s.cookies)}).")

    # -- report ----------------------------------------------------------- #

    def run_report_xlsx(self, report_id: str, plan_id: str, start: str, end: str,
                        poll_seconds: int = 3, max_wait: int = 900) -> bytes:
        """start/end as dd/mm/yyyy. Returns the exported .xlsx bytes."""
        ui_ref = self.host + "/accurate/"
        init = self._post_json(self.host + "/accurate/report/init-report-input.do",
                               {"id": report_id, "planId": plan_id, "_dsi": self.dsi}, referer=ui_ref)["d"]
        report_input = json.loads(init["reportInput"])
        report_input.setdefault("param", {})
        report_input["param"]["startDate"] = start
        report_input["param"]["endDate"] = end

        ex = self._post_json(self.report_host + "/accurate/report/bg-execute-report.do", {
            "id": report_id, "planId": plan_id, "reportInput": json.dumps(report_input),
            "cacheId": "", "pageIndex": "0", "_usi": self.usi, "_dsi": self.dsi,
        }, referer=ui_ref)
        bg_pid = ex.get("b") or (ex.get("d") or {}).get("bgPid")
        if not bg_pid:
            raise AccurateError(f"bg-execute-report returned no background id (keys: {sorted(ex)})")

        deadline = time.time() + max_wait
        while True:
            pr = self._post_json(self.report_host + "/accurate/company/bg-proc-response.do", {
                "bgPid": bg_pid, "keepCache": "true", "_usi": self.usi, "_dsi": self.dsi,
            }, referer=ui_ref)["d"]
            status = pr.get("status")
            if status == "FINISHED":
                resp = pr.get("response") or {}
                if not resp.get("s", True):
                    raise AccurateError(f"report failed: {resp.get('d')}")
                cache_id = resp.get("cacheId")
                break
            if status not in (None, "RUNNING", "PROCESSING", "QUEUE", "QUEUED", "WAITING"):
                raise AccurateError(f"report background job status {status!r}")
            if time.time() > deadline:
                raise AccurateError(f"report still {status!r} after {max_wait}s")
            time.sleep(poll_seconds)

        r = self.s.post(self.report_host + "/accurate/report/export-report.do", data={
            "_usi": self.usi, "_dsi": self.dsi, "cacheId": cache_id, "exportType": "xls", "name": "",
        }, headers={"Referer": ui_ref, "Origin": self.host}, timeout=max(self.timeout, 300))
        r.raise_for_status()
        if not r.content.startswith(b"PK"):
            raise AccurateError(f"export did not return an xlsx ({r.headers.get('content-type')}, {len(r.content)} bytes)")
        return r.content


def final_path(r: requests.Response) -> str:
    u = urlsplit(r.url)
    return f"{u.netloc}{u.path}"
