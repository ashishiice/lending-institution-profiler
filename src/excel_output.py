"""
Lending Institution Profiler — Excel Output Generator
======================================================
Generates a structured Excel workbook with:
1. Lender Matrix  — instrument × institution lending profile
2. Sector Focus   — where each lender is heading next
3. Raw Data       — scraped / manual data backup

Author: Bolt (AI Assistant for Ashish Prakash)
"""

import os
import json
from datetime import datetime
from pathlib import Path

import pandas as pd
import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo

# ── Paths ──────────────────────────────────────────────────────────────────────
PROJECT_ROOT = Path(__file__).parent.parent
DATA_RAW     = PROJECT_ROOT / "data" / "raw"
DATA_PROC    = PROJECT_ROOT / "data" / "processed"
OUTPUTS      = PROJECT_ROOT / "outputs"
OUTPUTS.mkdir(parents=True, exist_ok=True)

# ── Instrument Taxonomy ────────────────────────────────────────────────────────
INSTRUMENTS = [
    "INR Term Loans",
    "Working Capital Lines / CC / OD",
    "External Commercial Borrowings (ECB)",
    "Direct Assignment (DA)",
    "PTC / Securitisation",
    "Trade Finance (LC / BG)",
    "Corporate Loans / CPS / CCPS",
    "Infrastructure Finance",
    "G-Sec / SDL Purchases",
    "SME / MSME Lending",
    "Agriculture / Agri Finance",
    "NBFC / HFC Refinance",
]

# Lending product flags per instrument
PRODUCT_FLAGS = {
    "INR Term Loans":             "TL",
    "Working Capital Lines / CC / OD": "WC",
    "External Commercial Borrowings (ECB)": "ECB",
    "Direct Assignment (DA)":     "DA",
    "PTC / Securitisation":       "PTC",
    "Trade Finance (LC / BG)":    "TF",
    "Corporate Loans / CPS / CCPS": "CPS",
    "Infrastructure Finance":     "INF",
    "G-Sec / SDL Purchases":      "G-SEC",
    "SME / MSME Lending":         "MSME",
    "Agriculture / Agri Finance": "AGR",
    "NBFC / HFC Refinance":      "REF",
}

# ── Colour Palette ─────────────────────────────────────────────────────────────
C_DARK_BLUE   = "1F3864"  # Header bg
C_MID_BLUE    = "2E75B6"  # Sub-header
C_LIGHT_BLUE  = "D6E4F0"  # Alternating rows
C_GOLD        = "C9A227"  # Highlight / focus
C_GREEN       = "375623"  # Positive indicator
C_AMBER       = "ED7D31"  # Medium priority
C_RED         = "C00000"  # Weak / not offered
C_WHITE       = "FFFFFF"
C_LIGHT_GREY  = "F2F2F2"
C_DARK_GREY   = "595959"

# Instrument column colours (12 columns → 12 shades)
INST_COLOURS = [
    "DAEEF3", "E2F0D9", "FCE4D6", "FFF2CC",
    "E2EFDA", "F4CCCC", "D9D2E9", "FEE599",
    "DDEBF7", "F2DCDB", "E2EFDA", "D9E1F2",
]


def _fill(hex_colour: str) -> PatternFill:
    return PatternFill(fill_type="solid", fgColor=hex_colour)


def _font(bold: bool = False, colour: str = "000000",
          size: int = 10, name: str = "Calibri") -> Font:
    return Font(bold=bold, color=colour, size=size, name=name)


def _align(h: str = "left", v: str = "center", wrap: bool = False) -> Alignment:
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def _thin_border() -> Border:
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)


def _style_cell(cell, fill=None, font=None, align=None, border=None):
    if fill:   cell.fill       = fill
    if font:   cell.font       = font
    if align:  cell.alignment  = align
    if border: cell.border     = border


# ── Load Data ─────────────────────────────────────────────────────────────────

def load_latest_raw() -> list:
    """Find and load the most recent consolidated raw JSON."""
    files = sorted(DATA_RAW.glob("consolidated_*.json"), reverse=True)
    if not files:
        return []
    with open(files[0], encoding="utf-8") as f:
        return json.load(f)


# ── Sheet 1: Lender Matrix ─────────────────────────────────────────────────────

def _score(value) -> str:
    """Convert raw estimate to a human-readable availability score."""
    if value in ("excluded", "not offered", "N/A", "nil", 0, "0"):
        return "❌ N/A"
    if isinstance(value, (int, float)):
        if value > 0:
            return "✅ Strong"
    if isinstance(value, str):
        vl = value.lower()
        if any(k in vl for k in ["strong", "primary", "major", "large"]):
            return "✅ Strong"
        if any(k in vl for k in ["limited", "small", "niche"]):
            return "⚠️ Limited"
        if any(k in vl for k in ["growing", "focus", "target"]):
            return "🔶 Growing"
        if any(k in vl for k in ["excluded"]):
            return "❌ N/A"
    return "⚠️ Check"


def _notes_for_instrument(inst_row: dict, instrument_key: str) -> str:
    """Return the notes/amount text for a specific instrument."""
    # This relies on the scraped dict having instrument keys
    return inst_row.get("instruments", {}).get(instrument_key, "")


def build_lender_matrix(wb: openpyxl.Workbook, raw_data: list):

    ws = wb.create_sheet("Lender Matrix")
    ws.sheet_view.showGridLines = False

    # ── Title ─────────────────────────────────────────────────────────────────
    ws.merge_cells("A1:M1")
    title_cell = ws["A1"]
    title_cell.value = "LENDING INSTITUTION PROFILER — LENDER MATRIX"
    _style_cell(title_cell,
                fill=_fill(C_DARK_BLUE),
                font=_font(bold=True, colour=C_WHITE, size=14),
                align=_align("center"))

    ws.merge_cells("A2:M2")
    sub = ws["A2"]
    sub.value = f"Corporate lending instruments only — retail excluded  |  Generated: {datetime.today().strftime('%d %b %Y')}"
    _style_cell(sub,
                fill=_fill(C_MID_BLUE),
                font=_font(colour=C_WHITE, size=9),
                align=_align("center"))

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 18

    # ── Column headers (row 3) ────────────────────────────────────────────────
    col_headers = ["#", "Institution", "Type"] + INSTRUMENTS
    for c_idx, hdr in enumerate(col_headers, start=1):
        cell = ws.cell(row=3, column=c_idx, value=hdr)
        if c_idx <= 2:
            _style_cell(cell,
                        fill=_fill(C_DARK_BLUE),
                        font=_font(bold=True, colour=C_WHITE, size=10),
                        align=_align("center"))
        else:
            inst_col_idx = c_idx - 3  # 0-based index into INSTRUMENTS
            col_colour = INST_COLOURS[inst_col_idx % len(INST_COLOURS)]
            _style_cell(cell,
                        fill=_fill(col_colour),
                        font=_font(bold=True, colour=C_DARK_GREY, size=9),
                        align=_align("center", wrap=True))

    ws.row_dimensions[3].height = 40

    # ── Data rows ────────────────────────────────────────────────────────────
    for row_idx, inst in enumerate(raw_data, start=4):
        fill = _fill(C_LIGHT_BLUE) if row_idx % 2 == 0 else _fill(C_WHITE)

        # #
        _style_cell(ws.cell(row=row_idx, column=1, value=row_idx - 3),
                    fill=fill, font=_font(size=9), align=_align("center"))

        # Institution name
        _style_cell(ws.cell(row=row_idx, column=2, value=inst.get("institution", "")),
                    fill=fill, font=_font(bold=True, size=10), align=_align("left"))

        # Type badge
        _style_cell(ws.cell(row=row_idx, column=3, value=inst.get("type", "")),
                    fill=fill, font=_font(size=9, colour=C_DARK_GREY),
                    align=_align("center"))

        # Instrument scores
        for inst_idx, instrument in enumerate(INSTRUMENTS, start=4):
            raw_val = _notes_for_instrument(inst, instrument)
            score   = _score(raw_val)

            cell = ws.cell(row=row_idx, column=inst_idx, value=score)
            if "✅" in score:
                _style_cell(cell, fill=_fill("E2EFDA"),
                            font=_font(size=9, colour=C_GREEN, bold=True),
                            align=_align("center"))
            elif "🔶" in score:
                _style_cell(cell, fill=_fill("FFF2CC"),
                            font=_font(size=9, colour=C_AMBER, bold=True),
                            align=_align("center"))
            elif "❌" in score:
                _style_cell(cell, fill=_fill("FCE4D6"),
                            font=_font(size=9, colour=C_RED),
                            align=_align("center"))
            else:
                _style_cell(cell, fill=fill,
                            font=_font(size=9, colour=C_DARK_GREY),
                            align=_align("center"))

        ws.row_dimensions[row_idx].height = 22

    # ── Legend ────────────────────────────────────────────────────────────────
    legend_row = len(raw_data) + 5
    ws.merge_cells(f"A{legend_row}:M{legend_row}")
    leg = ws[f"A{legend_row}"]
    leg.value = "LEGEND   ✅ Strong = Primary product / major portfolio   |   🔶 Growing = Strategic focus area   |   ⚠️ Limited = Niche or small book   |   ❌ N/A = Not offered / out of scope"
    _style_cell(leg, fill=_fill(C_LIGHT_GREY),
                font=_font(size=8, colour=C_DARK_GREY),
                align=_align("left", wrap=True))
    ws.row_dimensions[legend_row].height = 20

    # ── Column widths ─────────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 14
    for i in range(4, 4 + len(INSTRUMENTS)):
        ws.column_dimensions[get_column_letter(i)].width = 20

    return ws


# ── Sheet 2: Sector Focus ──────────────────────────────────────────────────────

def build_sector_focus(wb: openpyxl.Workbook, raw_data: list):

    ws = wb.create_sheet("Sector Focus")
    ws.sheet_view.showGridLines = False

    # Title
    ws.merge_cells("A1:J1")
    title_cell = ws["A1"]
    title_cell.value = "LENDING INSTITUTION PROFILER — SECTOR FOCUS & STRATEGIC DIRECTION"
    _style_cell(title_cell,
                fill=_fill(C_DARK_BLUE),
                font=_font(bold=True, colour=C_WHITE, size=14),
                align=_align("center"))
    ws.row_dimensions[1].height = 28

    ws.merge_cells("A2:J2")
    sub = ws["A2"]
    sub.value = "Where each institution is placing their chips next — from annual reports, press releases, investor presentations"
    _style_cell(sub,
                fill=_fill(C_MID_BLUE),
                font=_font(colour=C_WHITE, size=9),
                align=_align("center"))
    ws.row_dimensions[2].height = 18

    # Headers
    headers = ["#", "Institution", "Type", "Top Sector Focus (FY25)",
               "Secondary Focus", "Tertiary Focus", "Products to Pitch",
               "Products to Avoid", "Notes / Colour", "Data Confidence"]
    for c_idx, hdr in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=c_idx, value=hdr)
        _style_cell(cell,
                    fill=_fill(C_DARK_BLUE),
                    font=_font(bold=True, colour=C_WHITE, size=9),
                    align=_align("center", wrap=True))
    ws.row_dimensions[3].height = 36

    # Sector list for lookup
    ALL_SECTORS = [
        "Infrastructure (Roads, Energy, Railways, Ports)",
        "Renewable / Green Energy (Solar, Wind)",
        "MSME / SME / Startup Financing",
        "Agriculture / Agri Infrastructure",
        "Affordable Housing / HFC",
        "Digital Infrastructure / Fintech",
        "Supply Chain Finance",
        "Export Finance / Trade Finance",
        "Healthcare / Pharma",
        "Manufacturing / Capex",
        "Real Estate / Commercial Real Estate",
        "NBFC / HFC Refinance",
        "Education / EdTech",
    ]

    for row_idx, inst in enumerate(raw_data, start=4):
        fill = _fill(C_LIGHT_BLUE) if row_idx % 2 == 0 else _fill(C_WHITE)

        _style_cell(ws.cell(row=row_idx, column=1, value=row_idx - 3),
                    fill=fill, font=_font(size=9), align=_align("center"))

        _style_cell(ws.cell(row=row_idx, column=2, value=inst.get("institution", "")),
                    fill=fill, font=_font(bold=True, size=10))

        _style_cell(ws.cell(row=row_idx, column=3, value=inst.get("type", "")),
                    fill=fill, font=_font(size=9), align=_align("center"))

        # Sector focus (from scraped data)
        sectors = inst.get("sector_focus_fy25", [])
        for col_off, sector in enumerate(sectors[:3], start=4):
            _style_cell(ws.cell(row=row_idx, column=col_off, value=sector),
                        fill=_fill("EBF3E8"),
                        font=_font(size=9),
                        align=_align("left", wrap=True))

        # Blank remaining sector columns
        for col_off in range(4 + len(sectors), 7):
            _style_cell(ws.cell(row=row_idx, column=col_off, value="—"),
                        fill=fill, font=_font(size=9, colour="AAAAAA"),
                        align=_align("center"))

        # Products to pitch (derived from instruments dict)
        strong_instr = [k for k, v in inst.get("instruments", {}).items()
                        if _score(v).startswith("✅")]
        pitch_text = "\n".join(f"• {p}" for p in strong_instr[:5])
        _style_cell(ws.cell(row=row_idx, column=7, value=pitch_text),
                    fill=_fill("E2EFDA"),
                    font=_font(size=9),
                    align=_align("left", wrap=True))

        # Products to avoid
        exclude = [k for k, v in inst.get("instruments", {}).items()
                   if _score(v).startswith("❌")]
        avoid_text = "\n".join(f"• {p}" for p in exclude) if exclude else "• ECB (typically lender)"
        _style_cell(ws.cell(row=row_idx, column=8, value=avoid_text),
                    fill=_fill("FCE4D6"),
                    font=_font(size=9),
                    align=_align("left", wrap=True))

        # Notes
        _style_cell(ws.cell(row=row_idx, column=9, value=inst.get("notes", "")),
                    fill=fill, font=_font(size=8, colour=C_DARK_GREY),
                    align=_align("left", wrap=True))

        # Confidence
        conf = inst.get("confidence", "medium")
        conf_cell = ws.cell(row=row_idx, column=10, value=conf)
        if conf == "high":
            _style_cell(conf_cell, fill=_fill("E2EFDA"),
                        font=_font(size=9, colour=C_GREEN, bold=True),
                        align=_align("center"))
        else:
            _style_cell(conf_cell, fill=_fill("FFF2CC"),
                        font=_font(size=9, colour=C_AMBER),
                        align=_align("center"))

        ws.row_dimensions[row_idx].height = 50

    # Column widths
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 24
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 28
    ws.column_dimensions["E"].width = 24
    ws.column_dimensions["F"].width = 22
    ws.column_dimensions["G"].width = 28
    ws.column_dimensions["H"].width = 26
    ws.column_dimensions["I"].width = 32
    ws.column_dimensions["J"].width = 14

    return ws


# ── Sheet 3: Raw Data ─────────────────────────────────────────────────────────

def build_raw_sheet(wb: openpyxl.Workbook, raw_data: list):
    ws = wb.create_sheet("Raw Data")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:B1")
    title = ws["A1"]
    title.value = "Raw Scraped / Manual Data — Edit Values Here"
    _style_cell(title,
                fill=_fill(C_DARK_GREY),
                font=_font(bold=True, colour=C_WHITE, size=12),
                align=_align("center"))
    ws.row_dimensions[1].height = 22

    ws.merge_cells("A2:B2")
    sub = ws["A2"]
    sub.value = "This sheet is the source of truth. Update manually when new annual reports / press releases are published."
    _style_cell(sub,
                fill=_fill(C_LIGHT_GREY),
                font=_font(size=8, colour=C_DARK_GREY),
                align=_align("left"))
    ws.row_dimensions[2].height = 16

    row = 4
    for inst in raw_data:
        ws.merge_cells(f"A{row}:B{row}")
        hdr = ws[f"A{row}"]
        hdr.value = f"{inst.get('institution', '')} ({inst.get('short_name', '')}) — {inst.get('type', '')}"
        _style_cell(hdr,
                    fill=_fill(C_MID_BLUE),
                    font=_font(bold=True, colour=C_WHITE, size=10),
                    align=_align("left"))
        ws.row_dimensions[row].height = 18
        row += 1

        # Total advances
        ws.cell(row=row, column=1, value="Total Advances (₹ Cr)").font = _font(bold=True, size=9)
        ws.cell(row=row, column=2, value=inst.get("total_advances", "")).font = _font(size=9)
        ws.row_dimensions[row].height = 16
        row += 1

        # Segments
        for seg, val in inst.get("segments", {}).items():
            ws.cell(row=row, column=1, value=f"  {seg}").font = _font(size=9)
            ws.cell(row=row, column=2, value=str(val)).font = _font(size=9)
            ws.row_dimensions[row].height = 15
            row += 1

        # Instruments
        ws.cell(row=row, column=1, value="LENDING INSTRUMENTS").font = _font(bold=True, size=9, colour=C_MID_BLUE)
        ws.row_dimensions[row].height = 16
        row += 1
        for instr, note in inst.get("instruments", {}).items():
            ws.cell(row=row, column=1, value=f"  {instr}").font = _font(size=9)
            ws.cell(row=row, column=2, value=str(note)).font = _font(size=9)
            ws.row_dimensions[row].height = 15
            row += 1

        # Sector focus
        sectors = inst.get("sector_focus_fy25", [])
        if sectors:
            ws.cell(row=row, column=1, value="SECTOR FOCUS FY25").font = _font(bold=True, size=9, colour=C_MID_BLUE)
            ws.row_dimensions[row].height = 16
            row += 1
            for s in sectors:
                ws.cell(row=row, column=1, value=f"  → {s}").font = _font(size=9)
                ws.row_dimensions[row].height = 15
                row += 1

        # Source
        ws.cell(row=row, column=1, value="Source").font = _font(size=8, colour=C_DARK_GREY)
        ws.cell(row=row, column=2, value=inst.get("source", "")).font = _font(size=8, colour=C_DARK_GREY)
        ws.row_dimensions[row].height = 15
        row += 2  # gap between institutions

    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 65
    return ws


# ── Sheet 4: Instrument Legend ─────────────────────────────────────────────────

def build_legend_sheet(wb: openpyxl.Workbook):
    ws = wb.create_sheet("Instrument Legend")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:D1")
    title = ws["A1"]
    title.value = "INSTRUMENT TAXONOMY — DEFINITIONS & PITCH NOTES"
    _style_cell(title,
                fill=_fill(C_DARK_BLUE),
                font=_font(bold=True, colour=C_WHITE, size=13),
                align=_align("center"))
    ws.row_dimensions[1].height = 26

    headers = ["Code", "Instrument", "What It Means", "When to Pitch"]
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=c, value=h)
        _style_cell(cell,
                    fill=_fill(C_MID_BLUE),
                    font=_font(bold=True, colour=C_WHITE, size=9),
                    align=_align("center"))
    ws.row_dimensions[2].height = 22

    definitions = [
        ("TL",    "INR Term Loans",
         "Plain vanilla term loans in Indian Rupees. Tenor typically 1-10 years. Fixed or floating rate.",
         "Best for: BoM, BoI, PNB, Union Bank, SIDBI. They have strong INR lending appetite and competitive rates."),
        ("WC",    "Working Capital Lines / CC / OD",
         "Cash credit, overdraft, working capital demand loan. Renewed annually. Secured by current assets.",
         "Best for: HDFC Bank, SBI, Canara Bank. High volume, relationship-driven."),
        ("ECB",   "External Commercial Borrowings",
         "Foreign currency borrowings by Indian entities. Governed by FEMA / RBI ECB guidelines. Tenor ≥3Y.",
         "NOT a borrowing product for most PSU banks — they lend in INR. Skip ECB pitch for PSU banks. Relevant for HDFC Bank, IDFC First, EXIM Bank."),
        ("DA",    "Direct Assignment",
         "Outright purchase of loan assets from an originating lender (NBFC/bank). Off-balance sheet for buyer.",
         "Best buyers: HDFC Bank, SBI, BoB, Union Bank. Strong DA desks. Pitch when you have NBFC assets to sell."),
        ("PTC",   "PTC / Securitisation",
         "Pass-through certificates — structured debt backed by pooled assets. Investors get principal + interest from pool cashflows.",
         "Best investors: HDFC Bank, SBI, BoB. Most active in senior PTC tranches. Pitch for senior tranche subscription."),
        ("TF",    "Trade Finance (LC / BG)",
         "Letters of Credit, Bank Guarantees, usance LC. Trade-related contingent exposures.",
         "Best for: EXIM Bank (exports), SBI, BoB (import trade). SIDBI also does supply chain finance."),
        ("CPS",   "Corporate Loans / CPS / CCPS",
         "Compulsorily Convertible Debentures, Optionally Convertible Debentures, corporate term loans.",
         "HDFC Bank, IDFC First, BoB have strong corporate finance teams. Check sector alignment."),
        ("INF",   "Infrastructure Finance",
         "Loans to roads, power, railways, ports, airports, social infrastructure. Long tenor (10-20Y).",
         "Best for: BoB (IBU), SBI (infrastructure), EXIM Bank (export-linked infra), NHB (social housing infra)."),
        ("G-SEC", "G-Sec / SDL Purchases",
         "Primary / secondary market purchases of Government Securities or State Development Loans.",
         "All banks invest. Not a direct pitch — more relevant for liquidity management discussions."),
        ("MSME",  "SME / MSME Lending",
         "Loans to Micro, Small, Medium Enterprises. PSL-qualified. Interest rates 10-20%.",
         "Best for: SIDBI (refinance + direct), SBI, PNB, Canara Bank, Central Bank. SIDBI is the REF financing partner."),
        ("AGR",   "Agriculture / Agri Finance",
         "Loans to farmers, agri-processing, warehouse receipts, Kisan Credit Cards.",
         "Best for: SBI, PNB, BoM, Bank of India. NHB also for agri-housing."),
        ("REF",   "NBFC / HFC Refinance",
         "Refinancing of existing NBFC / HFC loan books. Provides long-term liquidity to lenders.",
         "Best for: NHB (housing), SIDBI (MSME), EXIM Bank (export finance NBFCs). NHB is the #1 HFC refinancier."),
    ]

    for row_idx, (code, instr, definition, pitch) in enumerate(definitions, start=3):
        fill = _fill(C_LIGHT_BLUE) if row_idx % 2 == 0 else _fill(C_WHITE)

        c_cell = ws.cell(row=row_idx, column=1, value=code)
        _style_cell(c_cell, fill=_fill("DDEBF7"), font=_font(bold=True, size=9),
                    align=_align("center"))

        i_cell = ws.cell(row=row_idx, column=2, value=instr)
        _style_cell(i_cell, fill=fill, font=_font(bold=True, size=9),
                    align=_align("left"))

        d_cell = ws.cell(row=row_idx, column=3, value=definition)
        _style_cell(d_cell, fill=fill, font=_font(size=9),
                    align=_align("left", wrap=True))

        p_cell = ws.cell(row=row_idx, column=4, value=pitch)
        _style_cell(p_cell, fill=_fill("E2EFDA"), font=_font(size=9),
                    align=_align("left", wrap=True))

        ws.row_dimensions[row_idx].height = 48

    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 48
    ws.column_dimensions["D"].width = 55
    return ws


# ── Sheet 5: Institution Detail ───────────────────────────────────────────────

def build_institution_cards(wb: openpyxl.Workbook, raw_data: list):
    """One section per institution — all data consolidated."""

    ws = wb.create_sheet("Institution Detail")
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:G1")
    title = ws["A1"]
    title.value = "INSTITUTION DEEP-DIVE — ALL DATA PER INSTITUTION"
    _style_cell(title,
                fill=_fill(C_DARK_BLUE),
                font=_font(bold=True, colour=C_WHITE, size=13),
                align=_align("center"))
    ws.row_dimensions[1].height = 26

    row = 3
    for inst in raw_data:
        # Institution header
        ws.merge_cells(f"A{row}:G{row}")
        hdr = ws[f"A{row}"]
        hdr.value = f"{inst.get('institution', '')}  [{inst.get('short_name', '')}]  —  {inst.get('type', '')}"
        _style_cell(hdr,
                    fill=_fill(C_MID_BLUE),
                    font=_font(bold=True, colour=C_WHITE, size=11),
                    align=_align("left"))
        ws.row_dimensions[row].height = 22
        row += 1

        # Total advances
        ws.merge_cells(f"A{row}:B{row}")
        ws[f"A{row}"].value = "Total Advances (₹ Cr)"
        ws[f"A{row}"].font = _font(bold=True, size=9)

        ws.merge_cells(f"C{row}:G{row}")
        ws[f"C{row}"].value = inst.get("total_advances", "TBD")
        ws[f"C{row}"].font = _font(size=9)
        ws.row_dimensions[row].height = 16
        row += 1

        # Instrument matrix (compact)
        ws.cell(row=row, column=1, value="INSTRUMENT").font = _font(bold=True, size=9, colour=C_WHITE)
        ws.merge_cells(f"A{row}:A{row}")
        _style_cell(ws.cell(row=row, column=1, value="INSTRUMENT"),
                     fill=_fill(C_DARK_GREY))
        _style_cell(ws.cell(row=row, column=2, value="STATUS / NOTES"),
                     fill=_fill(C_DARK_GREY), font=_font(bold=True, colour=C_WHITE, size=9))
        _style_cell(ws.cell(row=row, column=3, value=""),
                     fill=_fill(C_DARK_GREY))
        _style_cell(ws.cell(row=row, column=4, value="SECTOR FOCUS"),
                     fill=_fill(C_DARK_GREY), font=_font(bold=True, colour=C_WHITE, size=9))
        ws.merge_cells(f"D{row}:G{row}")
        ws.row_dimensions[row].height = 18
        row += 1

        instruments = inst.get("instruments", {})
        sectors     = inst.get("sector_focus_fy25", [])
        for i_idx, (instr, note) in enumerate(instruments.items()):
            fill = _fill(C_LIGHT_BLUE) if i_idx % 2 == 0 else _fill(C_WHITE)
            score_cell = ws.cell(row=row, column=1, value=instr)
            _style_cell(score_cell, fill=fill, font=_font(size=9, bold=True))

            score_text = _score(note)
            val_cell = ws.cell(row=row, column=2, value=note)
            if "✅" in score_text:
                _style_cell(val_cell, fill=_fill("E2EFDA"), font=_font(size=9, colour=C_GREEN))
            elif "❌" in score_text:
                _style_cell(val_cell, fill=_fill("FCE4D6"), font=_font(size=9, colour=C_RED))
            else:
                _style_cell(val_cell, fill=fill, font=_font(size=9))

            # Sector focus (only in first row)
            if i_idx == 0:
                sector_text = "\n".join(f"• {s}" for s in sectors) if sectors else "—"
                ws.merge_cells(f"D{row}:G{row}")
                _style_cell(ws.cell(row=row, column=4, value=sector_text),
                            fill=_fill("EBF3E8"), font=_font(size=9),
                            align=_align("left", wrap=True))
            else:
                for c in range(3, 8):
                    _style_cell(ws.cell(row=row, column=c, value=""), fill=fill)

            ws.row_dimensions[row].height = 36
            row += 1

        # Notes
        ws.merge_cells(f"A{row}:B{row}")
        ws[f"A{row}"].value = "Notes"
        ws[f"A{row}"].font = _font(bold=True, size=9, colour=C_DARK_GREY)
        ws.merge_cells(f"C{row}:G{row}")
        ws[f"C{row}"].value = inst.get("notes", "")
        ws[f"C{row}"].font = _font(size=9, colour=C_DARK_GREY)
        ws.row_dimensions[row].height = 16
        row += 1

        # Source
        ws.merge_cells(f"A{row}:B{row}")
        ws[f"A{row}"].value = "Source"
        ws[f"A{row}"].font = _font(size=8, colour=C_DARK_GREY)
        ws.merge_cells(f"C{row}:G{row}")
        ws[f"C{row}"].value = inst.get("source", "")
        ws[f"C{row}"].font = _font(size=8, colour=C_DARK_GREY)
        ws.row_dimensions[row].height = 15
        row += 3  # gap between cards

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 12
    return ws


# ── Main Generator ─────────────────────────────────────────────────────────────

def generate_excel(raw_data: list, out_name: str = None) -> str:
    """Generate full Excel workbook and return path."""

    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Build all sheets
    build_lender_matrix(wb, raw_data)
    build_sector_focus(wb, raw_data)
    build_institution_cards(wb, raw_data)
    build_legend_sheet(wb)
    build_raw_sheet(wb, raw_data)

    # Save
    if out_name is None:
        out_name = f"Lender_Profiler_{datetime.today().strftime('%Y%m%d')}.xlsx"
    out_path = OUTPUTS / out_name
    wb.save(out_path)
    print(f"\n✅ Excel saved: {out_path}")
    return str(out_path)


if __name__ == "__main__":
    raw = load_latest_raw()
    if not raw:
        print("⚠️  No raw data found. Run scraper.py first.")
        print("   Falling back to sample data for demo...")
        # Use built-in sample data as fallback
        from scraper import INSTITUTIONS, InstitutionScraper, SBIScraper, HDFCBankScraper, Session
        session = Session()
        raw = []
        for inst in INSTITUTIONS:
            if inst["id"] == "sbi":
                s = SBIScraper(inst, session)
            elif inst["id"] == "hdfc_bank":
                s = HDFCBankScraper(inst, session)
            else:
                s = InstitutionScraper(inst, session)
            raw.append({
                "institution": inst["name"],
                "short_name":  inst["short"],
                "type":        inst["type"],
                "notes":       inst.get("notes", ""),
                **s.scrape_annual_report(),
                "press_releases": s.scrape_press_releases(),
                "sector_focus":  s.scrape_sector_focus(),
            })
    generate_excel(raw)
