"""Speichert Rechnungen per GitHub-API in einem (privaten) Repo und führt ein Journal."""
from __future__ import annotations

import base64
import csv
import io

import requests

API = "https://api.github.com"
JOURNAL = "rechnungen/journal.csv"
JOURNAL_COLS = ["rechnungs_nr", "rechnungsdatum", "positionen", "total_chf", "excel", "pdf",
                "erstellt", "final", "final_am", "daten"]


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

    def delete(self, path: str, message: str) -> None:
        _, sha = self.get(path)
        if not sha:
            return
        r = self.s.delete(self._url(path), json={"message": message, "sha": sha, "branch": self.branch})
        if r.status_code != 404:
            r.raise_for_status()

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
        """Kleinste freie Nummer ab START_NR. Entwürfe und finale Rechnungen behalten ihre Nummer,
        gelöschte Entwürfe hinterlassen keine Lücke (ihre Nummer wird wieder vergeben)."""
        belegt = {_nr(r) for r in rows}
        n = start
        while n in belegt:
            n += 1
        return n

    def naechste_nr(self, start: int) -> int:
        return self.naechste_nr_aus(self.journal()[0], start)

    def speichern(self, nr: str, name: str, xlsx: bytes, pdf: bytes, eintrag: dict,
                  daten: bytes | None = None) -> dict:
        rows, jsha = self.journal()
        alt = next((r for r in rows if r.get("rechnungs_nr") == str(nr)), None)
        if alt and ist_final(alt):
            raise RechnungFinal(f"Rechnung {nr} ist bereits final und wird nicht überschrieben.")
        self.ensure_branch()
        jahr = eintrag["rechnungsdatum"][-4:]
        p_x, p_p = f"rechnungen/{jahr}/{name}.xlsx", f"rechnungen/{jahr}/{name}.pdf"
        p_j = f"rechnungen/{jahr}/{name}.json" if daten is not None else ""
        dateien = [(p_x, xlsx), (p_p, pdf)] + ([(p_j, daten)] if daten is not None else [])
        for p, data in dateien:
            _, sha = self.get(p)
            self.put(p, data, f"Rechnung {nr}", sha)
        if alt:                                    # Entwurf überarbeitet → alte Dateien mit anderem Namen weg
            neu = {p_x, p_p, p_j}
            for k in ("excel", "pdf", "daten"):
                p = alt.get(k)
                if p and p not in neu:
                    try:
                        self.delete(p, f"Rechnung {nr}: alte Entwurfsdatei entfernt")
                    except Exception:
                        pass
        rows = [r for r in rows if r.get("rechnungs_nr") != str(nr)]
        rows.append({**eintrag, "excel": p_x, "pdf": p_p, "daten": p_j, "final": "nein", "final_am": ""})
        url = self._write_journal(rows, jsha, f"Journal: Rechnung {nr} (Entwurf)")
        return {"excel": p_x, "pdf": p_p, "journal_url": url, "rows": sorted(rows, key=_nr)}

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
        return rows

    def loeschen(self, nrs: list[str]) -> list[dict]:
        """Löscht Entwürfe (Dateien + Journal-Eintrag). Finale Rechnungen werden nicht gelöscht."""
        rows, sha = self.journal()
        weg = [r for r in rows if r.get("rechnungs_nr") in nrs]
        finals = [r["rechnungs_nr"] for r in weg if ist_final(r)]
        if finals:
            raise RechnungFinal(f"Rechnung {', '.join(finals)} ist final – zuerst den Final-Haken entfernen.")
        for r in weg:
            for p in (r.get("excel"), r.get("pdf"), r.get("daten")):
                if p:
                    self.delete(p, f"Rechnung {r['rechnungs_nr']} gelöscht")
        rows = [r for r in rows if r.get("rechnungs_nr") not in nrs]
        self._write_journal(rows, sha, f"Journal: Rechnung {', '.join(nrs)} gelöscht")
        return sorted(rows, key=_nr)
