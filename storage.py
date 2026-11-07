"""Speichert Rechnungen per GitHub-API in einem (privaten) Repo und führt ein Journal."""
from __future__ import annotations

import base64
import csv
import io

import requests

API = "https://api.github.com"
JOURNAL = "rechnungen/journal.csv"
JOURNAL_COLS = ["rechnungs_nr", "rechnungsdatum", "positionen", "total_chf", "excel", "pdf",
                "erstellt", "final", "final_am"]


class RechnungFinal(Exception):
    """Rechnung ist als final markiert und darf nicht überschrieben werden."""


def _nr(row: dict) -> int:
    v = str(row.get("rechnungs_nr", ""))
    return int(v) if v.isdigit() else 0


def ist_final(row: dict) -> bool:
    return str(row.get("final", "")).strip().lower() in ("ja", "true", "1", "x")


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
        """Legt den Daten-Branch an, falls er fehlt."""
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
        content = j.get("content")
        if not content and j.get("download_url"):          # Dateien > 1 MB
            raw = self.s.get(j["download_url"]); raw.raise_for_status()
            return raw.content, j["sha"]
        return base64.b64decode(content or ""), j["sha"]

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
        rows = list(csv.DictReader(io.StringIO(raw.decode("utf-8-sig"))))
        for r in rows:
            for c in JOURNAL_COLS:
                r.setdefault(c, "")
                if r[c] is None:
                    r[c] = ""
        return rows, sha

    def _write_journal(self, rows: list[dict], sha: str | None, msg: str) -> str:
        rows = sorted(rows, key=_nr)
        buf = io.StringIO()
        w = csv.DictWriter(buf, fieldnames=JOURNAL_COLS, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
        return self.put(JOURNAL, buf.getvalue().encode("utf-8-sig"), msg, sha)

    @staticmethod
    def naechste_nr_aus(rows: list[dict], start: int) -> int:
        """Nächste Nummer = höchste FINALE Nummer + 1. Entwürfe werden wiederverwendet/überschrieben."""
        finals = [_nr(r) for r in rows if ist_final(r)]
        return max(finals + [start - 1]) + 1

    def naechste_nr(self, start: int) -> int:
        return self.naechste_nr_aus(self.journal()[0], start)

    def speichern(self, nr: str, name: str, xlsx: bytes, pdf: bytes, eintrag: dict) -> dict:
        rows, jsha = self.journal()
        alt = next((r for r in rows if r.get("rechnungs_nr") == str(nr)), None)
        if alt and ist_final(alt):
            raise RechnungFinal(f"Rechnung {nr} ist bereits final und wird nicht überschrieben.")
        self.ensure_branch()
        jahr = eintrag["rechnungsdatum"][-4:]
        p_x, p_p = f"rechnungen/{jahr}/{name}.xlsx", f"rechnungen/{jahr}/{name}.pdf"
        for p, data in ((p_x, xlsx), (p_p, pdf)):
            _, sha = self.get(p)
            self.put(p, data, f"Rechnung {nr}", sha)
        rows = [r for r in rows if r.get("rechnungs_nr") != str(nr)]
        rows.append({**eintrag, "excel": p_x, "pdf": p_p, "final": "nein", "final_am": ""})
        url = self._write_journal(rows, jsha, f"Journal: Rechnung {nr} (Entwurf)")
        return {"excel": p_x, "pdf": p_p, "journal_url": url}

    def final_setzen(self, status: dict[str, bool], zeitstempel: str) -> int:
        """status: {rechnungs_nr: True/False}. Gibt Anzahl geänderter Einträge zurück."""
        rows, sha = self.journal()
        n = 0
        for r in rows:
            nr = r.get("rechnungs_nr")
            if nr in status and status[nr] != ist_final(r):
                r["final"] = "ja" if status[nr] else "nein"
                r["final_am"] = zeitstempel if status[nr] else ""
                n += 1
        if n:
            self._write_journal(rows, sha, f"Journal: Final-Status ({n} Änderung/en)")
        return n
