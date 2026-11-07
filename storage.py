"""Speichert Rechnungen per GitHub-API in einem (privaten) Repo und führt ein Journal."""
from __future__ import annotations

import base64
import csv
import io

import requests

API = "https://api.github.com"
JOURNAL = "rechnungen/journal.csv"
JOURNAL_COLS = ["rechnungs_nr", "rechnungsdatum", "positionen", "total_chf", "excel", "pdf", "erstellt"]


class GitHubStore:
    def __init__(self, token: str, repo: str, branch: str = "rechnungen"):
        self.repo, self.branch = repo, branch
        self.s = requests.Session()
        self.s.headers.update({"Authorization": f"Bearer {token}",
                               "Accept": "application/vnd.github+json",
                               "X-GitHub-Api-Version": "2022-11-28"})

    # -- Hilfen
    def _url(self, path: str) -> str:
        return f"{API}/repos/{self.repo}/contents/{path}"

    def ensure_branch(self) -> None:
        """Legt den Daten-Branch an, falls er fehlt (getrennt vom App-Branch → kein Neustart der App)."""
        r = self.s.get(f"{API}/repos/{self.repo}/branches/{self.branch}")
        if r.status_code == 200:
            return
        repo = self.s.get(f"{API}/repos/{self.repo}")
        repo.raise_for_status()
        default = repo.json()["default_branch"]
        ref = self.s.get(f"{API}/repos/{self.repo}/git/ref/heads/{default}")
        ref.raise_for_status()
        sha = ref.json()["object"]["sha"]
        self.s.post(f"{API}/repos/{self.repo}/git/refs",
                    json={"ref": f"refs/heads/{self.branch}", "sha": sha}).raise_for_status()

    def get(self, path: str) -> tuple[bytes | None, str | None]:
        r = self.s.get(self._url(path), params={"ref": self.branch})
        if r.status_code == 404:
            return None, None
        r.raise_for_status()
        j = r.json()
        return base64.b64decode(j["content"]), j["sha"]

    def put(self, path: str, content: bytes, message: str, sha: str | None = None) -> str:
        body = {"message": message, "content": base64.b64encode(content).decode(), "branch": self.branch}
        if sha:
            body["sha"] = sha
        r = self.s.put(self._url(path), json=body)
        r.raise_for_status()
        return r.json()["content"]["html_url"]

    # -- Journal
    def journal(self) -> tuple[list[dict], str | None]:
        raw, sha = self.get(JOURNAL)
        if raw is None:
            return [], None
        return list(csv.DictReader(io.StringIO(raw.decode("utf-8-sig")))), sha

    def naechste_nr(self, start: int) -> int:
        rows, _ = self.journal()
        nrs = [int(r["rechnungs_nr"]) for r in rows if r.get("rechnungs_nr", "").isdigit()]
        return max(nrs + [start - 1]) + 1

    def speichern(self, nr: str, name: str, xlsx: bytes, pdf: bytes, eintrag: dict) -> dict:
        self.ensure_branch()
        jahr = eintrag["rechnungsdatum"][-4:]
        p_x, p_p = f"rechnungen/{jahr}/{name}.xlsx", f"rechnungen/{jahr}/{name}.pdf"
        for p, data in ((p_x, xlsx), (p_p, pdf)):
            _, sha = self.get(p)                     # existiert schon → überschreiben (Korrektur)
            self.put(p, data, f"Rechnung {nr}", sha)
        rows, sha = self.journal()
        rows = [r for r in rows if r.get("rechnungs_nr") != str(nr)]
        rows.append({**eintrag, "excel": p_x, "pdf": p_p})
        rows.sort(key=lambda r: int(r["rechnungs_nr"]) if r["rechnungs_nr"].isdigit() else 0)
        buf = io.StringIO()
        w = csv.DictWriter(buf, fieldnames=JOURNAL_COLS, extrasaction="ignore")
        w.writeheader(); w.writerows(rows)
        url = self.put(JOURNAL, buf.getvalue().encode("utf-8-sig"), f"Journal: Rechnung {nr}", sha)
        return {"excel": p_x, "pdf": p_p, "journal_url": url}
