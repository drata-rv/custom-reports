#!/usr/bin/env python3
"""
Drata GRC Executive Summary Report Generator
=============================================
Pulls control and monitor data from the Drata Public API V1 and produces
a formatted .docx report that matches the GRC Executive Summary template.

Requirements (install via pip):
    pip install requests python-docx matplotlib

Basic usage:
    python report_generator.py --token <API_TOKEN> --quarter 1 --year 2026 --company "Suncoast"

Full usage:
    python report_generator.py ^
        --token  <API_TOKEN>       ^
        --company "Suncoast"       ^
        --framework "CIS 8"        ^
        --quarter 1                ^
        --year 2026                ^
        --cis-range "1-18"         ^
        --output "Report_Q1.docx"

Supported --framework values (case-sensitive):
    "CIS 8", "SOC 2", "ISO 27001", "ISO 27001:2022", "HIPAA", "PCI DSS",
    "PCI DSS v4", "GDPR", "CCPA", "NIST 800-53", "NIST CSF", "NIST CSF 2.0",
    "CMMC", "NIST 800-171", "FedRAMP", "HITRUST", "SOX ITGC", "DORA", ...
    (any tag returned by the Drata API is also accepted)
"""

from __future__ import annotations

import argparse
import os
import sys
import time
import tempfile
from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional

import requests

import matplotlib
matplotlib.use("Agg")          # Non-interactive backend — required on Windows without a display
import matplotlib.pyplot as plt

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


# ─────────────────────────────────────────────────────────────────────────────
# Global constants
# ─────────────────────────────────────────────────────────────────────────────

PAGE_SIZE       = 50        # confirmed safe limit for this API
REQUEST_TIMEOUT = 30        # seconds per HTTP request
MAX_RETRIES     = 3         # attempts before giving up on a request

# Drata Public API V1 — US endpoint (the only region supported by this script)
BASE_URL = "https://public-api.drata.com/public"

# Mapping from human-readable framework names (as used with --framework and
# as returned in control.frameworkTags[] response values) to the enum strings
# the API's ?frameworkTags= query parameter actually accepts.
# Source: GET /controls parameter schema in the V1 OpenAPI spec.
FRAMEWORK_TAG_MAP: Dict[str, str] = {
    # CIS
    "CIS 8":             "CIS8",
    # SOC
    "SOC 2":             "SOC_2",
    # ISO
    "ISO 27001":         "ISO27001",
    "ISO 27001:2022":    "ISO270012022",
    "ISO 27017:2015":    "ISO270172015",
    "ISO 27018:2019":    "ISO270182019",
    "ISO 27018:2025":    "ISO270182025",
    "ISO 27701":         "ISO27701",
    "ISO 27701:2025":    "ISO277012025",
    "ISO 42001:2023":    "ISO420012023",
    # PCI
    "PCI DSS":           "PCI",
    "PCI DSS v4":        "PCI4",
    "PCI DSS v4.0.1":    "PCI4",
    # NIST
    "NIST 800-53":       "NIST80053",
    "NIST CSF":          "NISTCSF",
    "NIST CSF 2.0":      "NISTCSF2",
    "NIST 800-171":      "NIST800171",
    "NIST 800-171 R3":   "NIST800171R3",
    "NIST AI":           "NISTAI",
    # Compliance
    "HIPAA":             "HIPAA",
    "GDPR":              "GDPR",
    "CCPA":              "CCPA",
    "CCPA 2026":         "CCPA2026",
    "CMMC":              "CMMC",
    "FFIEC":             "FFIEC",
    "FedRAMP":           "FEDRAMP",
    "FedRAMP 20x":       "FEDRAMP20X",
    "HITRUST":           "HITRUST",
    "SOX ITGC":          "SOX_ITGC",
    "COBIT":             "COBIT",
    "SCF":               "SCF",
    "CCM":               "CCM",
    "NIS2":              "NIS2",
    "DORA":              "DORA",
    "MS-SSPA":           "MSSSPA",
    "MS-SSPA 1.1":       "MSSSPA11",
    "Cyber Essentials":  "CYBER_ESSENTIALS",
    "Cyber Essentials 3.2": "CYBER_ESSENTIALS_32",
    "Essential Eight":   "ESSENTIAL_EIGHT",
    "NYDFS":             "NYDFS",
    "TISAX":             "TISAX",
    "CPS 230":           "CPS230",
    "Drata Essentials":  "DRATA_ESSENTIALS",
    "Custom":            "CUSTOM",
}

# Document styling — colours match the template's steel-blue headings
HEADING_COLOR = RGBColor(0x1B, 0x5E, 0x8A)
CAPTION_COLOR = RGBColor(0x55, 0x55, 0x55)
HEADING_FONT  = "Calibri"
BODY_FONT     = "Calibri"
HEADING_PT    = Pt(18)
BODY_PT       = Pt(11)
CAPTION_PT    = Pt(10)

# Pie-chart colour palettes (one per figure, ordered by legend entry)
CHART_COLORS: Dict[str, List[str]] = {
    # Figure 1  — Pass / Fail / Pending / Error
    "fig1": ["#2ECC71", "#E74C3C", "#F1C40F", "#95A5A6"],
    # Figure 2  — policy implementation levels
    "fig2": ["#1A5276", "#2980B9", "#27AE60", "#F39C12", "#E74C3C", "#9B59B6"],
    # Figure 3  — implementation fullness
    "fig3": ["#27AE60", "#2980B9", "#F39C12", "#E74C3C", "#BDC3C7"],
    # Figure 4  — automation level
    "fig4": ["#27AE60", "#2980B9", "#F1C40F", "#E67E22", "#E74C3C", "#9B59B6"],
}


# ─────────────────────────────────────────────────────────────────────────────
# Drata API client
# ─────────────────────────────────────────────────────────────────────────────

class DrataAPIError(Exception):
    """Raised for unrecoverable API problems (bad token, network failure, etc.)."""


class DrataClient:
    """
    Thin wrapper around the Drata Public API V1.

    All list endpoints are paginated automatically; callers receive a single
    flat list regardless of how many pages the API returns.
    """

    def __init__(
        self,
        token: str,
        base_url: str = BASE_URL,
        verify_ssl: bool = True,
        proxy: Optional[str] = None,
    ) -> None:
        self._base_url = base_url.rstrip("/")
        self._session = requests.Session()
        self._session.headers.update({
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
        })
        # Don't auto-detect proxy settings from the Windows registry / env vars.
        # On corporate Windows machines requests picks up IE/WinINet proxy settings
        # which can block outbound calls that work fine in PowerShell.
        # The Drata API is reachable directly; we don't need a proxy.
        self._session.trust_env = False
        # SSL verification — set False only when corporate proxy does SSL inspection
        self._session.verify = verify_ssl
        if not verify_ssl:
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        # Explicit proxy (http/https) — overrides env-var proxy settings
        if proxy:
            self._session.proxies = {"http": proxy, "https": proxy}

    # ── Low-level request ────────────────────────────────────────────────────

    def _get(self, endpoint: str, params: dict) -> dict:
        url = f"{self._base_url}/{endpoint}"
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                resp = self._session.get(url, params=params, timeout=REQUEST_TIMEOUT)
            except requests.exceptions.SSLError as exc:
                if attempt == MAX_RETRIES:
                    raise DrataAPIError(
                        f"SSL certificate verification failed connecting to {url}.\n"
                        f"  Detail: {exc}\n\n"
                        f"  This is common on corporate networks that inspect HTTPS traffic.\n"
                        f"  Re-run with --insecure to bypass SSL verification, or use\n"
                        f"  --proxy http://your-corporate-proxy:port to route through a proxy."
                    )
                time.sleep(attempt * 2.0)
                continue
            except requests.exceptions.ConnectionError as exc:
                if attempt == MAX_RETRIES:
                    raise DrataAPIError(
                        f"Cannot connect to {url}.\n"
                        f"  Detail: {exc}\n\n"
                        f"  Possible causes:\n"
                        f"  - Corporate firewall blocking outbound HTTPS\n"
                        f"  - SSL inspection proxy (try --insecure)\n"
                        f"  - Proxy required (try --proxy http://proxy:port)\n"
                        f"  - No internet access from this machine"
                    )
                time.sleep(attempt * 2.0)
                continue
            except requests.exceptions.Timeout:
                if attempt == MAX_RETRIES:
                    raise DrataAPIError(
                        f"Request to /{endpoint} timed out after {REQUEST_TIMEOUT}s."
                    )
                time.sleep(attempt * 2.0)
                continue

            if resp.status_code == 401:
                raise DrataAPIError(
                    "Authentication failed. Verify your --token value."
                )
            if resp.status_code == 403:
                raise DrataAPIError(
                    f"Access denied for /{endpoint}. "
                    "Check that the token has the required API permissions."
                )
            if resp.status_code == 429:
                # Rate-limited — back off and retry
                retry_after = float(resp.headers.get("Retry-After", attempt * 5))
                _progress(f"  Rate limited — waiting {retry_after:.0f}s...")
                time.sleep(retry_after)
                continue
            if resp.status_code >= 500:
                if attempt == MAX_RETRIES:
                    raise DrataAPIError(
                        f"Drata API returned a server error ({resp.status_code}). "
                        "Try again later."
                    )
                time.sleep(attempt * 2.0)
                continue

            resp.raise_for_status()
            return resp.json()

        return {}

    # ── Pagination helper ────────────────────────────────────────────────────

    def _paginate(
        self, endpoint: str, extra_params: Optional[Dict[str, str]] = None
    ) -> List[dict]:
        results: List[dict] = []
        page    = 1
        total   = 0
        base    = extra_params.copy() if extra_params else {}

        while True:
            params = {**base, "limit": PAGE_SIZE, "page": page}
            data   = self._get(endpoint, params)
            batch  = data.get("data", [])

            if page == 1:
                total = data.get("total", 0)

            results.extend(batch)
            _progress(f"  [{endpoint}] fetched {len(results)} / {total}")

            if not batch or len(results) >= total:
                break
            page += 1

        print()   # end the overwriting progress line
        return results

    # ── Public methods ───────────────────────────────────────────────────────

    def get_controls(self, framework_tag: Optional[str] = None) -> List[dict]:
        """
        Fetch all controls, optionally filtered server-side by framework.

        `framework_tag` should be the human-readable name (e.g. "CIS 8").
        It is looked up in FRAMEWORK_TAG_MAP to get the enum value the API
        expects (e.g. "CIS8").  If the tag isn't in the map the filter is
        skipped and all controls are fetched — the client-side filter in
        ReportData will still narrow them down correctly.
        """
        extra: Dict[str, str] = {}
        if framework_tag:
            api_enum = FRAMEWORK_TAG_MAP.get(framework_tag)
            if api_enum:
                extra["frameworkTags"] = api_enum
            else:
                _progress(
                    f"  Note: '{framework_tag}' not in FRAMEWORK_TAG_MAP — "
                    f"fetching all controls and filtering client-side."
                )
                print()
        return self._paginate("controls", extra_params=extra or None)

    def get_monitors(self) -> List[dict]:
        return self._paginate("monitors")

    def validate_token(self) -> None:
        """
        Check that the token works against the Drata API endpoint.
        Raises DrataAPIError with the specific failure reason if anything is wrong,
        so callers always see the real error rather than a generic message.
        """
        data = self._get("controls", {"limit": 1, "page": 1})
        if "data" not in data:
            raise DrataAPIError(
                f"Unexpected response from /public/controls "
                f"(got keys: {list(data.keys())}). "
                f"The API may have changed or the token may lack Controls read permission."
            )


def _progress(msg: str) -> None:
    """Overwrite the current terminal line with a progress message."""
    print(f"\r{msg:<70}", end="", flush=True)


# ─────────────────────────────────────────────────────────────────────────────
# Report data processor
# ─────────────────────────────────────────────────────────────────────────────

class ReportData:
    """
    Processes raw controls and monitors into every data point the report needs.

    Categorisation logic
    ────────────────────
    For each control that belongs to the requested framework:
      - passing  : every linked monitor returned PASSED
      - failing  : at least one linked monitor returned FAILED
      - erroring : monitors exist but none are FAILED; all are ERROR / unknown
      - pending  : no monitors are linked yet (not yet in active testing)

    Figure-3 and Figure-4 derivation
    ─────────────────────────────────
    Because the Drata API V1 does not expose CIS-specific implementation or
    automation level fields, we derive tiers from the boolean readiness signals
    that *are* available on every control object:

      isReady      — all evidence valid, tests passing, policies published
      isMonitored  — at least one automated test is linked and running
      hasEvidence  — manual evidence has been uploaded
      hasOwner     — a control owner has been assigned
    """

    def __init__(
        self,
        controls: List[dict],
        monitors: List[dict],
        framework_tag: str,
    ) -> None:
        self.framework_tag = framework_tag

        # ── Filter to the requested framework ───────────────────────────────
        self.fw_controls: List[dict] = [
            c for c in controls
            if framework_tag in c.get("frameworkTags", [])
        ]

        if not self.fw_controls:
            available = sorted(
                {t for c in controls for t in c.get("frameworkTags", [])}
            )
            raise ValueError(
                f"No controls found for framework tag '{framework_tag}'.\n"
                f"Framework tags available in this workspace: "
                f"{available if available else ['(none)']}"
            )

        # ── Build control_id → [monitor, …] lookup ───────────────────────────
        self._ctrl_mons: Dict[int, List[dict]] = defaultdict(list)
        for m in monitors:
            for ctrl in m.get("controls", []):
                self._ctrl_mons[ctrl["id"]].append(m)

        # ── Categorise every framework control ───────────────────────────────
        self.passing:  List[dict] = []
        self.failing:  List[dict] = []
        self.erroring: List[dict] = []
        self.pending:  List[dict] = []
        self._categorise()

    def _categorise(self) -> None:
        for ctrl in self.fw_controls:
            mons = self._ctrl_mons.get(ctrl["id"], [])
            if not mons:
                self.pending.append(ctrl)
                continue
            statuses = {m["checkResultStatus"] for m in mons}
            if "FAILED" in statuses:
                self.failing.append(ctrl)
            elif statuses <= {"PASSED"}:
                self.passing.append(ctrl)
            else:
                # Only ERROR (or other non-PASSED/FAILED) statuses present
                self.erroring.append(ctrl)

    # ── Figure data ──────────────────────────────────────────────────────────

    def fig1(self) -> Dict[str, int]:
        """Pass / Fail / Pending — sourced directly from monitor test results."""
        result: Dict[str, int] = {
            "Pass":                   len(self.passing),
            "Fail":                   len(self.failing),
            "Pending Implementation": len(self.pending),
        }
        if self.erroring:
            result["Error / Not Connected"] = len(self.erroring)
        return {k: v for k, v in result.items() if v > 0}

    def fig2(self) -> Dict[str, int]:
        """
        Policy implementation levels — derived from control readiness flags.

        Mapping (highest match wins, evaluated top-to-bottom):
          isReady AND hasEvidence  → Approved Written Policy
          hasEvidence              → Written Policy
          isMonitored              → Partially Written Policy
                                     (automated testing implies documented process)
          hasOwner                 → Informal Policy
                                     (ownership exists but no formal artefacts yet)
          (none)                   → No Policy
        """
        counts: Dict[str, int] = defaultdict(int)
        for c in self.fw_controls:
            if c["isReady"] and c["hasEvidence"]:
                counts["Approved Written Policy"] += 1
            elif c["hasEvidence"]:
                counts["Written Policy"] += 1
            elif c["isMonitored"]:
                counts["Partially Written Policy"] += 1
            elif c["hasOwner"]:
                counts["Informal Policy"] += 1
            else:
                counts["No Policy"] += 1
        return {k: v for k, v in counts.items() if v > 0}

    def fig3(self) -> Dict[str, int]:
        """
        Implementation fullness — tiered by readiness signals.

        Mapping (highest match wins):
          isReady      → Implemented on All Systems
          isMonitored  → Implemented on Most Systems
          hasEvidence  → Implemented on Some Systems
          hasOwner     → Parts of Control Implemented
          (none)       → Not Implemented
        """
        counts: Dict[str, int] = defaultdict(int)
        for c in self.fw_controls:
            if c["isReady"]:
                counts["Implemented on All Systems"] += 1
            elif c["isMonitored"]:
                counts["Implemented on Most Systems"] += 1
            elif c["hasEvidence"]:
                counts["Implemented on Some Systems"] += 1
            elif c["hasOwner"]:
                counts["Parts of Control Implemented"] += 1
            else:
                counts["Not Implemented"] += 1
        return {k: v for k, v in counts.items() if v > 0}

    def fig4(self) -> Dict[str, int]:
        """
        Automation level — derived from monitor presence and pass/fail ratios.

        Mapping:
          Monitors exist AND all PASSED            → Automated on All Systems
          Monitors exist AND some PASSED, some not → Automated on Most Systems
          Monitors exist AND none PASSED            → Automated on Some Systems
          No monitors AND hasOwner                 → Parts of Policy Automated
          No monitors AND not hasOwner             → Not Automated
        """
        counts: Dict[str, int] = defaultdict(int)
        for c in self.fw_controls:
            mons = self._ctrl_mons.get(c["id"], [])
            if mons:
                statuses = [m["checkResultStatus"] for m in mons]
                n_pass = statuses.count("PASSED")
                n_fail = statuses.count("FAILED")
                if n_fail == 0 and n_pass > 0:
                    counts["Automated on All Systems"] += 1
                elif n_pass > 0:
                    counts["Automated on Most Systems"] += 1
                else:
                    counts["Automated on Some Systems"] += 1
            elif c["hasOwner"]:
                counts["Parts of Policy Automated"] += 1
            else:
                counts["Not Automated"] += 1
        return {k: v for k, v in counts.items() if v > 0}

    # ── Narrative helpers ────────────────────────────────────────────────────

    def pending_safeguards(self) -> List[Dict]:
        """Ordered list of {code, name} for controls pending implementation."""
        return [
            {"code": c["code"], "name": c["name"]}
            for c in sorted(self.pending, key=lambda x: x.get("code") or "")
        ]

    def failed_safeguards(self) -> List[Dict]:
        """List of {code, name, reason} for controls with failing monitors."""
        items = []
        for c in sorted(self.failing, key=lambda x: x.get("code") or ""):
            reason = self._first_instance_field(
                c["id"], "failedTestDescription", failed_only=True
            )
            items.append({"code": c["code"], "name": c["name"], "reason": reason})
        return items

    def remediation_items(self) -> List[Dict]:
        """List of {code, name, recommendation} for controls with failing monitors."""
        items = []
        for c in sorted(self.failing, key=lambda x: x.get("code") or ""):
            rec = self._first_instance_field(
                c["id"], "remedyDescription", failed_only=True
            )
            items.append({
                "code": c["code"],
                "name": c["name"],
                "recommendation": rec,
            })
        return items

    def _first_instance_field(
        self, ctrl_id: int, field: str, failed_only: bool = False
    ) -> str:
        """
        Walk every monitor linked to ctrl_id and return the first non-empty
        value of `field` found inside any monitorInstance.
        When failed_only=True, only inspects monitors whose top-level
        checkResultStatus is FAILED.
        """
        for m in self._ctrl_mons.get(ctrl_id, []):
            if failed_only and m.get("checkResultStatus") != "FAILED":
                continue
            for inst in m.get("monitorInstances", []):
                value = inst.get(field, "")
                if value and isinstance(value, str):
                    return value.strip()
        return ""

    @property
    def total(self) -> int:
        return len(self.fw_controls)


# ─────────────────────────────────────────────────────────────────────────────
# Pie-chart builder
# ─────────────────────────────────────────────────────────────────────────────

def _pct(n: int, total: int) -> str:
    """Return n as a percentage of total, formatted to one decimal place."""
    return f"{n / total * 100:.1f}" if total else "0.0"


def build_pie_chart(
    data: Dict[str, int],
    title: str,
    colors: List[str],
    out_path: str,
) -> None:
    """
    Render a pie chart from `data` and save it as a PNG to `out_path`.

    Zero-value categories are omitted automatically.  If every category is
    zero a placeholder "No data" chart is written instead so the document
    always contains four figures.
    """
    active = {k: v for k, v in data.items() if v > 0}

    fig, ax = plt.subplots(figsize=(8, 5))
    fig.patch.set_facecolor("white")

    if not active:
        ax.text(
            0.5, 0.5, "No data available",
            ha="center", va="center", fontsize=13, color="#888888",
        )
        ax.axis("off")
    else:
        labels  = list(active.keys())
        values  = list(active.values())
        # Repeat palette if there are more slices than defined colours
        palette = (colors * ((len(labels) // max(len(colors), 1)) + 1))[: len(labels)]

        wedges, _, autotexts = ax.pie(
            values,
            colors=palette,
            autopct="%1.1f%%",
            startangle=140,
            pctdistance=0.78,
        )
        for at in autotexts:
            at.set_fontsize(8)
            at.set_color("white")

        ax.legend(
            wedges,
            [f"{lbl}  ({v})" for lbl, v in zip(labels, values)],
            loc="center left",
            bbox_to_anchor=(1.02, 0.5),
            fontsize=9,
            frameon=False,
        )

    ax.set_title(title, fontsize=11, fontweight="bold", pad=14)
    plt.tight_layout()
    plt.savefig(out_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)


# ─────────────────────────────────────────────────────────────────────────────
# Word-document builder
# ─────────────────────────────────────────────────────────────────────────────

class ReportBuilder:
    """
    Assembles the GRC Executive Summary .docx from a populated ReportData object.

    Document structure
    ──────────────────
    1. Executive Summary  — narrative + bullet stats
    2. Scope              — narrative + dash list of pending safeguards
    3. Findings           — narrative + dash list of failed safeguards
    4. Remediation        — narrative + dash list of recommendations
    5. Metrics            — four embedded pie-chart images with captions
    """

    def __init__(
        self,
        data: ReportData,
        company: str,
        quarter: int,
        year: int,
        cis_range: str,
    ) -> None:
        self._data      = data
        self._company   = company
        self._quarter   = quarter
        self._year      = year
        self._cis_range = cis_range
        self._doc       = Document()
        self._setup_page()

    # ── Page & style setup ───────────────────────────────────────────────────

    def _setup_page(self) -> None:
        sec = self._doc.sections[0]
        sec.page_height   = Inches(11)
        sec.page_width    = Inches(8.5)
        sec.left_margin   = Inches(1.25)
        sec.right_margin  = Inches(1.25)
        sec.top_margin    = Inches(1.0)
        sec.bottom_margin = Inches(1.0)

        # Set the document-wide default font so every added run inherits it
        normal = self._doc.styles["Normal"]
        normal.font.name = BODY_FONT
        normal.font.size = BODY_PT

    # ── Paragraph helpers ────────────────────────────────────────────────────

    def _heading(self, text: str) -> None:
        """Bold blue section heading (matches template style)."""
        para = self._doc.add_paragraph()
        para.paragraph_format.space_before = Pt(16)
        para.paragraph_format.space_after  = Pt(4)
        run = para.add_run(text)
        run.bold              = True
        run.font.name         = HEADING_FONT
        run.font.size         = HEADING_PT
        run.font.color.rgb    = HEADING_COLOR

    def _body(self, text: str, space_before: int = 0) -> None:
        """Regular body paragraph."""
        para = self._doc.add_paragraph()
        para.paragraph_format.space_before = Pt(space_before)
        para.paragraph_format.space_after  = Pt(4)
        run = para.add_run(text)
        run.font.name = BODY_FONT
        run.font.size = BODY_PT

    def _bullet(self, bold_text: str, rest: str) -> None:
        """
        Bullet point where `bold_text` is rendered bold and `rest` is normal.
        Uses a filled-circle Unicode bullet (●) instead of a Word list style
        so the output is identical across every Word version on Windows.
        """
        para = self._doc.add_paragraph()
        para.paragraph_format.left_indent        = Inches(0.35)
        para.paragraph_format.first_line_indent  = Inches(-0.25)
        para.paragraph_format.space_after        = Pt(3)

        rb = para.add_run("\u25CF  " + bold_text)   # ● + non-breaking spaces
        rb.bold       = True
        rb.font.name  = BODY_FONT
        rb.font.size  = BODY_PT

        rn = para.add_run(rest)
        rn.font.name  = BODY_FONT
        rn.font.size  = BODY_PT

    def _dash_item(self, text: str) -> None:
        """Indented list item prefixed with an en-dash (–)."""
        para = self._doc.add_paragraph()
        para.paragraph_format.left_indent       = Inches(0.35)
        para.paragraph_format.first_line_indent = Inches(-0.25)
        para.paragraph_format.space_after       = Pt(2)
        run = para.add_run(f"\u2013  {text}")    # en-dash
        run.font.name = BODY_FONT
        run.font.size = BODY_PT

    def _caption(self, text: str) -> None:
        """Italic figure caption in a slightly smaller, grey font."""
        para = self._doc.add_paragraph()
        para.paragraph_format.space_before = Pt(2)
        para.paragraph_format.space_after  = Pt(14)
        run = para.add_run(text)
        run.italic         = True
        run.font.name      = BODY_FONT
        run.font.size      = CAPTION_PT
        run.font.color.rgb = CAPTION_COLOR

    def _embed_chart(self, image_path: str) -> None:
        """Insert a chart PNG at 5.5 inches wide (fits within 6-inch text area)."""
        self._doc.add_picture(image_path, width=Inches(5.5))
        # python-docx centres pictures by default; keep left-aligned to match template
        self._doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.LEFT

    # ── Section writers ──────────────────────────────────────────────────────

    def _write_executive_summary(
        self, q: str, total: int,
        n_pass: int, n_fail: int, n_pend: int, n_err: int,
    ) -> None:
        self._heading("Executive Summary:")
        self._body(
            f"During {q} {self._year}, the Information Security Governance, Risk, and Compliance "
            f"(GRC) team completed control effectiveness testing for {self._data.framework_tag} "
            f"Controls {self._cis_range} ({total} related safeguards). The purpose of this testing "
            f"was to confirm whether key security safeguards are in place, working as intended, "
            f"and aligned with {self._company}\u2019s security objectives."
        )
        self._body(f"Summary of results ({total} safeguards tested):", space_before=6)
        self._bullet(
            f"{n_pass} passed ({_pct(n_pass, total)}%)",
            " \u2014 Implemented and operating effectively.",
        )
        self._bullet(
            f"{n_pend} pending implementation ({_pct(n_pend, total)}%)",
            " \u2014 Not yet in place; will be tested once implemented.",
        )
        self._bullet(
            f"{n_fail} failed ({_pct(n_fail, total)}%)",
            " \u2014 Gaps identified that require remediation to meet "
            "the intended safeguard requirements.",
        )
        if n_err:
            self._bullet(
                f"{n_err} error / not connected ({_pct(n_err, total)}%)",
                " \u2014 Monitoring connection issue; results are inconclusive.",
            )

    def _write_scope(self) -> None:
        # Calculate the *next* quarter for the re-assessment date
        next_q = self._quarter % 4 + 1
        next_y = self._year if self._quarter < 4 else self._year + 1

        self._heading("Scope:")
        self._body(
            f"The following safeguards are still in the implementation phase and will be "
            f"reassessed either in Q{next_q} {next_y} or once custom control automation tests "
            f"within the Drata platform are available for use, if these safeguards qualify "
            f"for automation."
        )
        items = self._data.pending_safeguards()
        if items:
            for item in items:
                label = (
                    f"{item['code']} \u2013 {item['name']}"
                    if item["code"]
                    else item["name"]
                )
                self._dash_item(label)
        else:
            self._body(
                "All safeguards are currently in active testing \u2014 "
                "no items are pending implementation."
            )

    def _write_findings(self, n_fail: int) -> None:
        s    = "s" if n_fail != 1 else ""
        were = "were" if n_fail != 1 else "was"

        self._heading("Findings:")
        self._body(
            f"Out of the total number of safeguards tested for compliance and effectiveness, "
            f"{n_fail} safeguard{s} did not meet the required standards and {were} marked as "
            f"failed. Below is a summary of the failed safeguards and the reasons for their "
            f"failure."
        )
        items = self._data.failed_safeguards()
        if items:
            for item in items:
                label = f"{item['code']} \u2013 {item['name']}"
                if item.get("reason"):
                    label += f": {item['reason']}"
                self._dash_item(label)
        else:
            self._body("No safeguards failed during this testing period.")

    def _write_remediation(self) -> None:
        self._heading("Remediation and Recommendations:")
        self._body(
            f"For safeguards that failed testing or were identified as needing improvement, "
            f"timely remediation is essential\u2014not only to strengthen "
            f"{self._company}\u2019s security posture but also to maintain alignment with "
            f"CIS standards. Below are targeted recommendations for addressing the failed "
            f"safeguards:"
        )
        items = self._data.remediation_items()
        if items:
            for item in items:
                label = f"{item['code']} \u2013 {item['name']}"
                if item.get("recommendation"):
                    label += f": {item['recommendation']}"
                self._dash_item(label)
        else:
            self._body("No remediation actions required at this time.")

    def _write_metrics(self, q: str, total: int, tmpdir: str) -> None:
        self._heading("Metrics:")

        charts = [
            (
                self._data.fig1(),
                f"{self._year} {q} Control Testing Results",
                CHART_COLORS["fig1"],
                f"Figure 1: Control Status after {q} testing. "
                f"{total} sub-controls were reviewed.",
            ),
            (
                self._data.fig2(),
                "Control Policy Implementation",
                CHART_COLORS["fig2"],
                "Figure 2: Controls based on Written Policy Implementation",
            ),
            (
                self._data.fig3(),
                f"{self._year} {q} Control Implementation",
                CHART_COLORS["fig3"],
                "Figure 3: Controls based on Fullness/Effectiveness of their Implementation",
            ),
            (
                self._data.fig4(),
                f"{self._year} {q} Control Automation Implementation",
                CHART_COLORS["fig4"],
                "Figure 4: Control based on levels of Automation Implementation",
            ),
        ]

        for idx, (chart_data, title, colors, caption) in enumerate(charts, 1):
            img_path = os.path.join(tmpdir, f"figure_{idx}.png")
            build_pie_chart(chart_data, title, colors, img_path)
            self._embed_chart(img_path)
            self._caption(caption)

    # ── Public build method ──────────────────────────────────────────────────

    def build(self, output_path: str) -> None:
        """Write the complete report to `output_path`."""
        d       = self._data
        q       = f"Q{self._quarter}"
        total   = d.total
        n_pass  = len(d.passing)
        n_fail  = len(d.failing)
        n_pend  = len(d.pending)
        n_err   = len(d.erroring)

        self._write_executive_summary(q, total, n_pass, n_fail, n_pend, n_err)
        self._write_scope()
        self._write_findings(n_fail)
        self._write_remediation()

        # Charts are written to a temp folder and embedded before it is cleaned up.
        # python-docx reads each PNG immediately on add_picture(), so the files
        # do not need to survive past the end of _write_metrics().
        with tempfile.TemporaryDirectory() as tmpdir:
            self._write_metrics(q, total, tmpdir)

        self._doc.save(output_path)


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────

def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="report_generator.py",
        description="Generate a Drata GRC Executive Summary report (.docx)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples (Windows):\n"
            "  python report_generator.py --token TOKEN --company \"Suncoast\" "
            "--quarter 1 --year 2026\n\n"
            "  python report_generator.py --token TOKEN --framework \"SOC 2\" "
            "--company \"Acme\" --quarter 2 --cis-range \"1-18\" "
            "--output \"Report_Q2.docx\"\n"
        ),
    )
    parser.add_argument(
        "--token", required=True,
        help="Drata API bearer token",
    )
    parser.add_argument(
        "--company", default="Your Organization",
        help="Company name used throughout the narrative text (default: 'Your Organization')",
    )
    parser.add_argument(
        "--framework", default="CIS 8",
        help="Framework tag to report on — must match exactly what Drata stores "
             "(default: 'CIS 8')",
    )
    parser.add_argument(
        "--quarter", type=int, required=True, choices=[1, 2, 3, 4],
        help="Report quarter: 1, 2, 3, or 4",
    )
    parser.add_argument(
        "--year", type=int, default=datetime.now().year,
        help="Report year (default: current year)",
    )
    parser.add_argument(
        "--cis-range", dest="cis_range", default=None,
        help="Control range label for the narrative, e.g. '1-18'. "
             "If omitted, the first and last control codes are used.",
    )
    parser.add_argument(
        "--output", default=None,
        help="Output .docx path. Default: "
             "GRC_Summary_<Company>_Q<N>_<Year>.docx in the current directory.",
    )
    parser.add_argument(
        "--insecure", action="store_true", default=False,
        help=(
            "Disable SSL certificate verification. Use this when running on a "
            "corporate network that performs SSL inspection (the proxy presents "
            "its own certificate which Python does not trust by default)."
        ),
    )
    parser.add_argument(
        "--proxy", default=None, metavar="URL",
        help=(
            "HTTP/HTTPS proxy URL to route API calls through, e.g. "
            "http://proxy.corp.example.com:8080. "
            "Only needed if your network requires an explicit proxy."
        ),
    )
    return parser


def main() -> None:
    args   = _build_parser().parse_args()
    output = args.output or (
        f"GRC_Summary_{args.company.replace(' ', '_')}_Q{args.quarter}_{args.year}.docx"
    )

    print("=" * 62)
    print("  Drata GRC Executive Summary Report Generator")
    print("=" * 62)
    print(f"  Company   : {args.company}")
    print(f"  Framework : {args.framework}")
    print(f"  Period    : Q{args.quarter} {args.year}")
    print(f"  Output    : {output}")
    print(f"  API       : {BASE_URL}")
    print()

    # ── Validate token ───────────────────────────────────────────────────────
    if args.insecure:
        print("  WARNING   : SSL verification disabled (--insecure)")
    if args.proxy:
        print(f"  Proxy     : {args.proxy}")
    print()

    client = DrataClient(
        args.token,
        verify_ssl=not args.insecure,
        proxy=args.proxy,
    )
    print("Validating API token...", end=" ", flush=True)
    try:
        client.validate_token()
        print("OK")
    except DrataAPIError as exc:
        print("FAILED")
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    # ── Fetch data ───────────────────────────────────────────────────────────
    print("\nFetching data from Drata API:")
    try:
        controls = client.get_controls(framework_tag=args.framework)
        print(f"  Controls fetched : {len(controls)}")
        monitors = client.get_monitors()
        print(f"  Monitors fetched : {len(monitors)}")
    except DrataAPIError as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    # ── Process ──────────────────────────────────────────────────────────────
    print("\nProcessing...")
    try:
        data = ReportData(controls, monitors, args.framework)
    except ValueError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)

    # Derive a control-range label for the narrative if the user did not supply one
    cis_range = args.cis_range
    if not cis_range:
        codes = sorted(
            c["code"] for c in data.fw_controls if c.get("code")
        )
        if len(codes) >= 2:
            cis_range = f"{codes[0]}\u2013{codes[-1]}"
        elif codes:
            cis_range = codes[0]
        else:
            cis_range = "All Controls"

    print(f"  {args.framework} controls total : {data.total}")
    print(f"  Passing                      : {len(data.passing)}")
    print(f"  Failing                      : {len(data.failing)}")
    print(f"  Pending (no monitors yet)    : {len(data.pending)}")
    if data.erroring:
        print(f"  Error / disconnected         : {len(data.erroring)}")

    # ── Build report ─────────────────────────────────────────────────────────
    print("\nBuilding report (generating charts & assembling document)...")
    builder = ReportBuilder(
        data=data,
        company=args.company,
        quarter=args.quarter,
        year=args.year,
        cis_range=cis_range,
    )
    builder.build(output)

    resolved = Path(output).resolve()
    print(f"\nDone.  Report saved to:\n  {resolved}\n")


if __name__ == "__main__":
    main()
