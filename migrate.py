#!/usr/bin/env python3
"""
================================================================================
Enterprise HR Data Migration & Reconciliation Tool
================================================================================
Migrates employee assignment and mobility data from source.xlsx into the
EXISTING Destination.xlsx workbook (Sheet1), strictly preserving formatting,
column widths, dropdowns, and header rows 1-3.

Produces an audit report workbook: migration_report.xlsx with 6 sheets:
  1. Summary: Key metrics, counts, and potential match count with renames
  2. Exact_Matches: All migrated fields, match types, and sample values
  3. Not_Matched_Source: Unmapped source columns with rename suggestions & confidence
  4. Not_Matched_Destination: Unmapped destination columns with closest source candidates
  5. Duplicates_And_Conflicts: Duplicate codes and their exact resolution
  6. Data_Issues: Suspicious values, country code formats, dropdown mismatches

CLI Arguments:
  --source         Path to source Excel file (default: source.xlsx)
  --destination    Path to destination Excel file (default: Destination.xlsx)
  --report         Path to report Excel file (default: migration_report.xlsx)
  --backup         Create a backup copy of Destination.xlsx before modifying
  --dry-run        Generate migration_report.xlsx without altering Destination.xlsx
================================================================================
"""

import os
import sys
import shutil
import re
import argparse
import datetime
import difflib
from typing import Dict, List, Tuple, Any, Optional, Set
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from dateutil import parser as date_parser

# ==============================================================================
# CONFIGURATION DICTIONARY
# ==============================================================================

MIGRATION_CONFIG: Dict[str, Any] = {
    # 1. Explicit renames (Source System Code / Header -> Destination System Code)
    "RENAMES": {
        # Home Company Segment Levels 1..4
        "ASSIGNMENT!FROM_COMPANY_LEVEL1": "ASSIGNMENT!FROM_COMPANY_SEGMENT_LEVEL_1",
        "ASSIGNMENT!FROM_COMPANY_LEVEL2": "ASSIGNMENT!FROM_COMPANY_SEGMENT_LEVEL_2",
        "ASSIGNMENT!FROM_COMPANY_LEVEL3": "ASSIGNMENT!FROM_COMPANY_SEGMENT_LEVEL_3",
        "ASSIGNMENT!FROM_COMPANY_LEVEL4": "ASSIGNMENT!FROM_COMPANY_SEGMENT_LEVEL_4",
        # Host Company Segment Levels 1..4
        "ASSIGNMENT!TO_COMPANY_LEVEL1": "ASSIGNMENT!TO_COMPANY_SEGMENT_LEVEL_1",
        "ASSIGNMENT!TO_COMPANY_LEVEL2": "ASSIGNMENT!TO_COMPANY_SEGMENT_LEVEL_2",
        "ASSIGNMENT!TO_COMPANY_LEVEL3": "ASSIGNMENT!TO_COMPANY_SEGMENT_LEVEL_3",
        "ASSIGNMENT!TO_COMPANY_LEVEL4": "ASSIGNMENT!TO_COMPANY_SEGMENT_LEVEL_4",
        # Group name
        "ASSIGNMENT!GROUP": "ASSIGNMENT!GROUP_NAME",
        # Employee Number
        "EMPLOYEE!SYSTEM_NUMBER": "Employee Number",
    },

    # 2. Explicit skip list (Source columns with no destination equivalent)
    "EXPLICIT_SKIPS": {
        "UT_GOVERNANCE_QUESTIONNAIRE2!PURPOSE",
        "ASSIGNMENT!OFFER_SALARY",
        "ASSIGNMENT!OFFER_SALARY_CURRENCY",
    },

    # 3. Command column defaults for Destination (filled only if blank)
    "COMMAND_DEFAULTS": {
        "A": "INSERT",                   # Col A (Action Command)
        "Assignment Command": "INSERT",  # Col B (Assignment Command)
        "Assignment Number": "NEW",      # Col D (Assignment Number)
    },

    # 4. Fuzzy matching threshold for suggested renames in the report
    "FUZZY_THRESHOLD": 0.75,

    # 5. Data row detection settings
    "HEADER_ROWS_COUNT": 4,  # Rows 1-4 are headers in source.xlsx; data starts on Row 5
}


# ==============================================================================
# STRING & VALUE HELPERS
# ==============================================================================

def normalize_text(text: Any) -> str:
    """Normalize text for exact comparison ignoring case, spaces, and punctuation."""
    if text is None:
        return ""
    return re.sub(r'[^a-zA-Z0-9]', '', str(text)).lower()


def is_blank_value(value: Any) -> bool:
    """Check if a cell value is empty, None, or literal 'None'."""
    if value is None:
        return True
    if isinstance(value, str):
        cleaned = value.strip()
        return cleaned == "" or cleaned.lower() == "none"
    return False


def is_comment_cell(code_str: Any) -> bool:
    """Check if a cell in source row 2 is an instructional comment rather than a code."""
    if not code_str:
        return False
    s = str(code_str).strip()
    if len(s) > 60 or "\n" in s:
        return True
    l = s.lower()
    return l.startswith("create ") or l.startswith("will need ") or l.startswith("new ") or l.startswith("add to ")


def format_us_date(value: Any) -> Tuple[Optional[str], Optional[str]]:
    """
    Format a date into MM/DD/YYYY format.
    Returns (formatted_date_string, error_message).
    """
    if is_blank_value(value):
        return None, None
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime("%m/%d/%Y"), None
    val_str = str(value).strip()
    try:
        dt = date_parser.parse(val_str)
        return dt.strftime("%m/%d/%Y"), None
    except Exception as e:
        return val_str, f"Unparseable date '{val_str}': {e}"


def calculate_similarity(s1: str, s2: str) -> float:
    """Calculate similarity ratio between two strings using raw and normalized forms."""
    if not s1 or not s2:
        return 0.0
    r1 = difflib.SequenceMatcher(None, s1.lower(), s2.lower()).ratio()
    n1, n2 = normalize_text(s1), normalize_text(s2)
    r2 = difflib.SequenceMatcher(None, n1, n2).ratio() if n1 and n2 else 0.0
    return max(r1, r2)


# ==============================================================================
# MAIN MIGRATOR CLASS
# ==============================================================================

class DataMigrator:
    def __init__(
        self,
        source_path: str,
        destination_path: str,
        report_path: str,
        backup: bool = False,
        dry_run: bool = False,
    ):
        self.source_path = os.path.abspath(source_path)
        self.destination_path = os.path.abspath(destination_path)
        self.report_path = os.path.abspath(report_path)
        self.backup = backup
        self.dry_run = dry_run

        # Parsed metadata
        self.dest_cols: List[Dict[str, Any]] = []
        self.dest_by_code: Dict[str, Dict[str, Any]] = {}
        self.dest_by_norm_name: Dict[str, Dict[str, Any]] = {}

        self.source_cols: List[Dict[str, Any]] = []
        self.source_data_rows: List[int] = []

        # Mapping structures
        self.mapped_source_to_dest: Dict[int, Dict[str, Any]] = {}  # sc -> dest_dict & match_info
        self.mapped_dest_cols: Dict[int, Dict[str, Any]] = {}       # dc -> source_dict & match_info

        # Reports & Audit Logs
        self.exact_matches: List[Dict[str, Any]] = []
        self.unmatched_source: List[Dict[str, Any]] = []
        self.unmatched_dest: List[Dict[str, Any]] = []
        self.duplicates_log: List[Dict[str, Any]] = []
        self.data_issues: List[Dict[str, Any]] = []

        # Metrics
        self.metrics = {
            "source_columns": 0,
            "destination_columns": 0,
            "exact_code_matches": 0,
            "renamed_matches": 0,
            "name_fallback_matches": 0,
            "duplicate_disambiguations": 0,
            "command_defaults": 0,
            "total_migrated_columns": 0,
            "unmatched_source_count": 0,
            "unmatched_destination_count": 0,
            "rows_migrated": 0,
            "potential_matches_with_renames": 0,
        }

    def validate_inputs(self) -> None:
        """Validate input files exist and are valid workbooks."""
        if not os.path.isfile(self.source_path):
            raise FileNotFoundError(f"Source file not found: {self.source_path}")
        if not os.path.isfile(self.destination_path):
            raise FileNotFoundError(f"Destination template file not found: {self.destination_path}")

    def run(self) -> None:
        """Execute the full migration workflow."""
        self.validate_inputs()

        # Step 1: Parse Destination headers
        print(f"Reading destination template: {self.destination_path}")
        wb_dest_read = openpyxl.load_workbook(self.destination_path, data_only=True)
        ws_dest_read = wb_dest_read["Sheet1"]
        self._parse_destination_headers(ws_dest_read)

        # Step 2: Parse Source headers & data
        print(f"Reading source workbook: {self.source_path}")
        wb_source = openpyxl.load_workbook(self.source_path, data_only=True)
        ws_source = wb_source["Sheet1"]
        self._parse_source_file(ws_source)

        # Step 3: Execute priority mapping logic
        self._execute_mapping_logic(ws_source)

        # Step 4: Scan for data issues & suspicious values
        self._detect_data_issues(ws_source)

        # Step 5: Compute fuzzy suggestions for unmapped columns
        self._compute_fuzzy_suggestions()

        # Step 6: Perform write into Destination.xlsx if not dry-run
        if not self.dry_run:
            self._write_migrated_data()
        else:
            print("\n[DRY RUN MODE] Destination.xlsx was NOT modified.")

        # Step 7: Generate comprehensive report workbook (migration_report.xlsx)
        self._generate_report_workbook()

        # Step 8: Print console summary
        self.print_summary()

    def _parse_destination_headers(self, ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
        """Parse rows 1-3 of Destination.xlsx."""
        for c in range(1, ws.max_column + 1):
            inst = ws.cell(1, c).value
            name = ws.cell(2, c).value
            code = ws.cell(3, c).value
            code_str = str(code).strip() if code is not None else ""
            name_str = str(name).strip() if name is not None else ""
            let = get_column_letter(c)
            is_us_date = bool(inst and "us date format" in str(inst).lower())

            col_info = {
                "col_idx": c,
                "col_letter": let,
                "inst": inst,
                "name": name_str,
                "code": code_str,
                "norm_name": normalize_text(name_str),
                "is_us_date": is_us_date,
            }
            self.dest_cols.append(col_info)
            if code_str:
                self.dest_by_code[code_str] = col_info
            if col_info["norm_name"]:
                self.dest_by_norm_name[col_info["norm_name"]] = col_info

        self.metrics["destination_columns"] = len(self.dest_cols)

    def _parse_source_file(self, ws: openpyxl.worksheet.worksheet.Worksheet) -> None:
        """Parse source headers (rows 1-4) and identify data rows (rows 5+)."""
        for c in range(1, ws.max_column + 1):
            r1 = ws.cell(1, c).value
            r2 = ws.cell(2, c).value
            r3 = ws.cell(3, c).value
            r4 = ws.cell(4, c).value
            code_str = str(r2).strip() if r2 is not None else ""
            ep_str = str(r3).strip() if r3 is not None else ""
            hdr_str = str(r4).strip() if r4 is not None else ""

            col_info = {
                "col_idx": c,
                "col_letter": get_column_letter(c),
                "inst": r1,
                "code": code_str,
                "ep": ep_str,
                "hdr": hdr_str,
                "is_comment": is_comment_cell(code_str),
            }
            self.source_cols.append(col_info)

        self.metrics["source_columns"] = len(self.source_cols)

        # Identify data rows (row 5 onward)
        start_row = 5
        for r in range(start_row, ws.max_row + 1):
            non_empty = [
                c for c in range(1, ws.max_column + 1)
                if not is_blank_value(ws.cell(r, c).value)
            ]
            if non_empty:
                self.source_data_rows.append(r)

        self.metrics["rows_migrated"] = len(self.source_data_rows)
        print(f"Found {len(self.source_data_rows)} active data row(s) in source (rows 5+).")

    def _execute_mapping_logic(self, ws_source: openpyxl.worksheet.worksheet.Worksheet) -> None:
        """
        Execute mapping in strict priority order:
          Priority 1: Primary Exact System Field Code (Source Row 2 == Dest Row 3)
          Priority 2: Explicit Renames & Legacy Topia ID
          Priority 3: Duplicate Codes Disambiguation (using Source Row 4 header text)
          Priority 4: Name Fallback (Source Row 3 / Row 4 exact normalized == Dest Row 2)
        """
        # --- PASS 1: High Priority (Exact Code, Renames, Disambiguated Duplicates) ---
        for sc_info in self.source_cols:
            sc = sc_info["col_idx"]
            code_str = sc_info["code"]
            hdr_str = sc_info["hdr"]
            ep_str = sc_info["ep"]
            comm = sc_info["is_comment"]

            if sc == 1:
                # Column A holds row labels only
                continue

            # Check Legacy Topia ID (Col 24 / Assignment ID) -> Dest Col E ('Assignment Identifier')
            if sc == 24 or (not code_str and hdr_str == "Assignment ID"):
                d = self.dest_by_code.get("Assignment Identifier")
                if d and d["col_idx"] not in self.mapped_dest_cols:
                    match_info = {
                        "dest": d,
                        "match_type": "EXPLICIT_RENAME",
                        "reason": "Mapped legacy Assignment ID (Topia ID) to Assignment Identifier (Col E)"
                    }
                    self.mapped_source_to_dest[sc] = match_info
                    self.mapped_dest_cols[d["col_idx"]] = {"source": sc_info, "info": match_info}
                    self.metrics["renamed_matches"] += 1
                continue

            # Check Explicit Skips
            if code_str in MIGRATION_CONFIG["EXPLICIT_SKIPS"]:
                continue

            # --- Duplicate Code Disambiguation by Source Header (Rule 5) ---
            # 1. Base Pay (CF_MONEY001 x2: Currency vs Amount)
            if code_str == "ASSIGNMENT_COMPENSATION_SALARY_CF!CF_MONEY001":
                if "currency" in hdr_str.lower():
                    d = self.dest_by_code.get("ASSIGNMENT_COMPENSATION_SALARY!ANNUAL_BASE_PAY_CURRENCY")
                    if d and d["col_idx"] not in self.mapped_dest_cols:
                        match_info = {"dest": d, "match_type": "DUPLICATE_DISAMBIGUATION", "reason": "Disambiguated by header 'Currency' -> ANNUAL_BASE_PAY_CURRENCY (Col 111)"}
                        self.mapped_source_to_dest[sc] = match_info
                        self.mapped_dest_cols[d["col_idx"]] = {"source": sc_info, "info": match_info}
                        self.duplicates_log.append({"code": code_str, "sc": sc_info, "dest": d, "res": "MAPPED", "notes": match_info["reason"]})
                        self.metrics["duplicate_disambiguations"] += 1
                elif "amount" in hdr_str.lower():
                    d = self.dest_by_code.get("ASSIGNMENT_COMPENSATION_SALARY!ANNUAL_BASE_PAY_AMOUNT")
                    if d and d["col_idx"] not in self.mapped_dest_cols:
                        match_info = {"dest": d, "match_type": "DUPLICATE_DISAMBIGUATION", "reason": "Disambiguated by header 'Amount' -> ANNUAL_BASE_PAY_AMOUNT (Col 112)"}
                        self.mapped_source_to_dest[sc] = match_info
                        self.mapped_dest_cols[d["col_idx"]] = {"source": sc_info, "info": match_info}
                        self.duplicates_log.append({"code": code_str, "sc": sc_info, "dest": d, "res": "MAPPED", "notes": match_info["reason"]})
                        self.metrics["duplicate_disambiguations"] += 1
                continue

            # 2. TO_TITLE (x2: Host Job Title vs Job Grade)
            if code_str == "ASSIGNMENT!TO_TITLE":
                if "grade" in ep_str.lower() or "level" in hdr_str.lower():
                    d = self.dest_by_code.get("ASSIGNMENT!TO_SALARY_GRADE")
                    if d and d["col_idx"] not in self.mapped_dest_cols:
                        match_info = {"dest": d, "match_type": "DUPLICATE_DISAMBIGUATION", "reason": "Disambiguated 2nd TO_TITLE ('Job Grade') -> ASSIGNMENT!TO_SALARY_GRADE (Col 67)"}
                        self.mapped_source_to_dest[sc] = match_info
                        self.mapped_dest_cols[d["col_idx"]] = {"source": sc_info, "info": match_info}
                        self.duplicates_log.append({"code": code_str, "sc": sc_info, "dest": d, "res": "MAPPED", "notes": match_info["reason"]})
                        self.metrics["duplicate_disambiguations"] += 1
                else:
                    d = self.dest_by_code.get("ASSIGNMENT!TO_TITLE")
                    if d and d["col_idx"] not in self.mapped_dest_cols:
                        match_info = {"dest": d, "match_type": "EXACT_CODE", "reason": "Primary TO_TITLE ('Host Job Title') -> ASSIGNMENT!TO_TITLE (Col 66)"}
                        self.mapped_source_to_dest[sc] = match_info
                        self.mapped_dest_cols[d["col_idx"]] = {"source": sc_info, "info": match_info}
                        self.duplicates_log.append({"code": code_str, "sc": sc_info, "dest": d, "res": "MAPPED", "notes": match_info["reason"]})
                        self.metrics["exact_code_matches"] += 1
                continue

            # 3. Benefits (EMPLOYEE_ELECTED_BENEFITS x2: Currency vs Amount)
            if code_str == "UT_CALCULATION_INPUT!EMPLOYEE_ELECTED_BENEFITS":
                if "currency" in hdr_str.lower():
                    d = self.dest_by_code.get("UT_CALCULATION_INPUT!EMPLOYEE_ELECTED_BENEFITS_CURRENCY")
                    if d and d["col_idx"] not in self.mapped_dest_cols:
                        match_info = {"dest": d, "match_type": "DUPLICATE_DISAMBIGUATION", "reason": "Disambiguated by header 'Currency' -> EMPLOYEE_ELECTED_BENEFITS_CURRENCY (Col 116)"}
                        self.mapped_source_to_dest[sc] = match_info
                        self.mapped_dest_cols[d["col_idx"]] = {"source": sc_info, "info": match_info}
                        self.duplicates_log.append({"code": code_str, "sc": sc_info, "dest": d, "res": "MAPPED", "notes": match_info["reason"]})
                        self.metrics["duplicate_disambiguations"] += 1
                elif "amount" in hdr_str.lower():
                    d = self.dest_by_code.get("UT_CALCULATION_INPUT!EMPLOYEE_ELECTED_BENEFITS")
                    if d and d["col_idx"] not in self.mapped_dest_cols:
                        match_info = {"dest": d, "match_type": "DUPLICATE_DISAMBIGUATION", "reason": "Disambiguated by header 'Amount' -> EMPLOYEE_ELECTED_BENEFITS (Col 117)"}
                        self.mapped_source_to_dest[sc] = match_info
                        self.mapped_dest_cols[d["col_idx"]] = {"source": sc_info, "info": match_info}
                        self.duplicates_log.append({"code": code_str, "sc": sc_info, "dest": d, "res": "MAPPED", "notes": match_info["reason"]})
                        self.metrics["duplicate_disambiguations"] += 1
                continue

            # 4. Mailing Address Location Types (x4: Work Home matches Col 92 'Work Home Address Type')
            if code_str == "ASSIGNMENT_MAILING_ADDRESS!LOCATION_TYPE":
                if "work home" in hdr_str.lower():
                    d = self.dest_by_code.get("ASSIGNMENT_MAILING_ADDRESS!LOCATION_TYPE")
                    if d and d["col_idx"] not in self.mapped_dest_cols:
                        match_info = {"dest": d, "match_type": "DUPLICATE_DISAMBIGUATION", "reason": "Disambiguated by header 'Work Home Address' -> Dest Col 92 'Work Home Address Type'"}
                        self.mapped_source_to_dest[sc] = match_info
                        self.mapped_dest_cols[d["col_idx"]] = {"source": sc_info, "info": match_info}
                        self.duplicates_log.append({"code": code_str, "sc": sc_info, "dest": d, "res": "MAPPED", "notes": match_info["reason"]})
                        self.metrics["duplicate_disambiguations"] += 1
                else:
                    self.duplicates_log.append({"code": code_str, "sc": sc_info, "dest": None, "res": "SKIPPED", "notes": f"Location type '{hdr_str}' has no dedicated column in destination"})
                continue

            # 5. Employee Contact Number Address (x2: Phone vs Mobile)
            if code_str == "ASSIGNMENT_EMPLOYEE_CONTACT!NUMBER_ADDRESS":
                if "mobile" not in hdr_str.lower():
                    d = self.dest_by_code.get("ASSIGNMENT_EMPLOYEE_CONTACT!NUMBER_ADDRESS")
                    if d and d["col_idx"] not in self.mapped_dest_cols:
                        match_info = {"dest": d, "match_type": "DUPLICATE_DISAMBIGUATION", "reason": "Primary Phone Number -> Dest Col 105 NUMBER_ADDRESS"}
                        self.mapped_source_to_dest[sc] = match_info
                        self.mapped_dest_cols[d["col_idx"]] = {"source": sc_info, "info": match_info}
                        self.duplicates_log.append({"code": code_str, "sc": sc_info, "dest": d, "res": "MAPPED", "notes": match_info["reason"]})
                        self.metrics["duplicate_disambiguations"] += 1
                else:
                    self.duplicates_log.append({"code": code_str, "sc": sc_info, "dest": None, "res": "SKIPPED", "notes": "Secondary Mobile Phone Number skipped in favor of primary Phone Number"})
                continue

            # Explicit Renames
            if code_str in MIGRATION_CONFIG["RENAMES"]:
                dest_code = MIGRATION_CONFIG["RENAMES"][code_str]
                d = self.dest_by_code.get(dest_code)
                if d and d["col_idx"] not in self.mapped_dest_cols:
                    match_info = {"dest": d, "match_type": "EXPLICIT_RENAME", "reason": f"Renamed from {code_str} to {dest_code}"}
                    self.mapped_source_to_dest[sc] = match_info
                    self.mapped_dest_cols[d["col_idx"]] = {"source": sc_info, "info": match_info}
                    self.metrics["renamed_matches"] += 1
                continue

            # Primary Exact Code Match
            if code_str and not comm and code_str in self.dest_by_code:
                d = self.dest_by_code[code_str]
                if d["col_idx"] not in self.mapped_dest_cols:
                    match_info = {"dest": d, "match_type": "EXACT_CODE", "reason": "Exact system field code match"}
                    self.mapped_source_to_dest[sc] = match_info
                    self.mapped_dest_cols[d["col_idx"]] = {"source": sc_info, "info": match_info}
                    self.metrics["exact_code_matches"] += 1

        # --- PASS 2: Name Fallback (Source Row 3 / Row 4 normalized match on Dest Row 2) ---
        for sc_info in self.source_cols:
            sc = sc_info["col_idx"]
            if sc in self.mapped_source_to_dest or sc == 1:
                continue

            code_str = sc_info["code"]
            hdr_str = sc_info["hdr"]
            ep_str = sc_info["ep"]
            comm = sc_info["is_comment"]

            if not code_str or comm:
                # Compare source row 3 or row 4 with destination friendly name
                m3 = self.dest_by_norm_name.get(normalize_text(ep_str)) if ep_str else None
                m4 = self.dest_by_norm_name.get(normalize_text(hdr_str)) if hdr_str else None
                target = m3 or m4
                if target:
                    dc = target["col_idx"]
                    if dc not in self.mapped_dest_cols:
                        which = "Row 3 friendly name" if target == m3 else "Row 4 header name"
                        matched_name = ep_str if target == m3 else hdr_str
                        match_info = {
                            "dest": target,
                            "match_type": "NAME_FALLBACK",
                            "reason": f"Exact normalized match on {which} '{matched_name}' to destination '{target['name']}'"
                        }
                        self.mapped_source_to_dest[sc] = match_info
                        self.mapped_dest_cols[dc] = {"source": sc_info, "info": match_info}
                        self.metrics["name_fallback_matches"] += 1

        self.metrics["total_migrated_columns"] = len(self.mapped_source_to_dest)

        # Collect exact match details for report
        for sc, info in sorted(self.mapped_source_to_dest.items()):
            sc_info = [s for s in self.source_cols if s["col_idx"] == sc][0]
            d = info["dest"]
            # Samples from row 5 and row 6
            v1 = ws_source.cell(5, sc).value if len(self.source_data_rows) >= 1 else None
            v2 = ws_source.cell(6, sc).value if len(self.source_data_rows) >= 2 else None
            self.exact_matches.append({
                "source_col_letter": sc_info["col_letter"],
                "source_col_idx": sc,
                "source_hdr": sc_info["hdr"],
                "source_code": sc_info["code"],
                "dest_col_letter": d["col_letter"],
                "dest_col_idx": d["col_idx"],
                "dest_name": d["name"],
                "dest_code": d["code"],
                "match_type": info["match_type"],
                "reason": info["reason"],
                "sample_v1": v1,
                "sample_v2": v2,
            })

    def _detect_data_issues(self, ws_source: openpyxl.worksheet.worksheet.Worksheet) -> None:
        """Scan migrated data rows for suspicious values, malformed text, or dropdown mismatches."""
        for r_idx, src_row in enumerate(self.source_data_rows, start=1):
            emp_id = ws_source.cell(src_row, 5).value  # Col E: EMPLOYEE!SYSTEM_NUMBER
            emp_name = ws_source.cell(src_row, 1).value  # Col A: Employee Full Name

            for sc, info in self.mapped_source_to_dest.items():
                sc_info = [s for s in self.source_cols if s["col_idx"] == sc][0]
                d = info["dest"]
                val = ws_source.cell(src_row, sc).value

                if is_blank_value(val):
                    continue

                val_str = str(val).strip()

                # Issue 1: Concatenated / corrupted text like 'Not ProvidedHome'
                if "home" in val_str.lower() and "not provided" in val_str.lower():
                    self.data_issues.append({
                        "row_num": 3 + r_idx,
                        "source_row": src_row,
                        "emp_id": emp_id,
                        "emp_name": emp_name,
                        "source_col": sc_info["col_letter"],
                        "source_hdr": sc_info["hdr"],
                        "dest_col": d["col_letter"],
                        "dest_name": d["name"],
                        "issue_type": "CORRUPTED_TEXT",
                        "value": val_str,
                        "desc": "Contains corrupted/concatenated text 'Home' appended to 'Not Provided'. Should be cleaned to 'Not Provided'."
                    })

                # Issue 2: Country name instead of 3-letter ISO country code
                if d["inst"] and "country code" in str(d["inst"]).lower():
                    # Country codes should be 3 uppercase letters (e.g. USA, IND, EGY, JPN)
                    if len(val_str) > 3:
                        self.data_issues.append({
                            "row_num": 3 + r_idx,
                            "source_row": src_row,
                            "emp_id": emp_id,
                            "emp_name": emp_name,
                            "source_col": sc_info["col_letter"],
                            "source_hdr": sc_info["hdr"],
                            "dest_col": d["col_letter"],
                            "dest_name": d["name"],
                            "issue_type": "COUNTRY_CODE_FORMAT",
                            "value": val_str,
                            "desc": f"Full country name '{val_str}' provided instead of 3-letter ISO Country Code per destination instruction '{d['inst']}'."
                        })

                # Issue 3: Potential dropdown mismatch
                if d["code"] == "EMPLOYEE!TYPE" and val_str.lower() in ["user", "employee"]:
                    if val_str != "Employee":
                        self.data_issues.append({
                            "row_num": 3 + r_idx,
                            "source_row": src_row,
                            "emp_id": emp_id,
                            "emp_name": emp_name,
                            "source_col": sc_info["col_letter"],
                            "source_hdr": sc_info["hdr"],
                            "dest_col": d["col_letter"],
                            "dest_name": d["name"],
                            "issue_type": "DROPDOWN_CASING",
                            "value": val_str,
                            "desc": f"Person Status '{val_str}' does not match standard capitalized dropdown value 'Employee'."
                        })

                if d["code"] == "EMPLOYEE!EMPLOYMENT_STATUS" and val_str.lower() == "employed":
                    self.data_issues.append({
                        "row_num": 3 + r_idx,
                        "source_row": src_row,
                        "emp_id": emp_id,
                        "emp_name": emp_name,
                        "source_col": sc_info["col_letter"],
                        "source_hdr": sc_info["hdr"],
                        "dest_col": d["col_letter"],
                        "dest_name": d["name"],
                        "issue_type": "DROPDOWN_MISMATCH",
                        "value": val_str,
                        "desc": "Employment status 'Employed' should map to destination dropdown value 'Active'."
                    })

                # Issue 4: Date parsing errors
                if d["is_us_date"]:
                    _, err = format_us_date(val)
                    if err:
                        self.data_issues.append({
                            "row_num": 3 + r_idx,
                            "source_row": src_row,
                            "emp_id": emp_id,
                            "emp_name": emp_name,
                            "source_col": sc_info["col_letter"],
                            "source_hdr": sc_info["hdr"],
                            "dest_col": d["col_letter"],
                            "dest_name": d["name"],
                            "issue_type": "UNPARSEABLE_DATE",
                            "value": val_str,
                            "desc": err
                        })

    def _compute_fuzzy_suggestions(self) -> None:
        """Find closest destination candidates for unmatched source and destination columns."""
        threshold = MIGRATION_CONFIG["FUZZY_THRESHOLD"]
        suggested_dest_claimed: Set[int] = set(self.mapped_dest_cols.keys())

        # Unmatched Source Columns
        for sc_info in self.source_cols:
            sc = sc_info["col_idx"]
            if sc in self.mapped_source_to_dest:
                continue

            reason = "No destination equivalent found"
            if sc == 1:
                reason = "Column A holds row labels only"
            elif sc_info["code"] in MIGRATION_CONFIG["EXPLICIT_SKIPS"]:
                reason = "Explicitly skipped per mapping rules (no destination equivalent)"
            elif sc_info["is_comment"]:
                reason = f"Row 2 contains instructional note instead of code: {sc_info['code'][:35]}..."
            elif not sc_info["code"]:
                reason = "No system field code specified in Row 2"

            # Check duplicate skips
            for dup in self.duplicates_log:
                if dup["sc"]["col_idx"] == sc and dup["res"] == "SKIPPED":
                    reason = dup["notes"]

            # Compute closest destination candidate among currently unmapped destination columns
            s_candidates = [x for x in [sc_info["hdr"], sc_info["ep"], sc_info["code"]] if x and not is_comment_cell(x)]
            best_dest = None
            best_score = 0.0

            for d in self.dest_cols:
                # Compare against unmapped destination columns
                d_candidates = [x for x in [d["name"], d["code"]] if x]
                for sc_val in s_candidates:
                    for dc_val in d_candidates:
                        score = calculate_similarity(str(sc_val), str(dc_val))
                        if score > best_score:
                            best_score = score
                            best_dest = d

            suggested_rename = "None"
            target_col_let = ""
            target_name = ""
            score_out = None

            if best_dest and best_score >= threshold and sc != 1:
                score_out = round(best_score, 2)
                target_col_let = best_dest["col_letter"]
                target_name = best_dest["code"] or best_dest["name"]
                source_ident = sc_info["code"] or sc_info["hdr"] or f"Col {sc_info['col_letter']}"
                suggested_rename = (
                    f'rename source column "{source_ident}" to "{target_name}" and it will match'
                )
                if best_dest["col_idx"] not in suggested_dest_claimed:
                    suggested_dest_claimed.add(best_dest["col_idx"])
                    self.metrics["potential_matches_with_renames"] += 1

            self.unmatched_source.append({
                "source_col_letter": sc_info["col_letter"],
                "source_col_idx": sc,
                "source_hdr": sc_info["hdr"],
                "source_ep": sc_info["ep"],
                "source_code": sc_info["code"],
                "reason": reason,
                "suggested_rename": suggested_rename,
                "target_col_letter": target_col_let,
                "target_name": target_name,
                "confidence_score": score_out,
            })

        self.metrics["unmatched_source_count"] = len(self.unmatched_source)

        # Unmatched Destination Columns
        for d in self.dest_cols:
            dc = d["col_idx"]
            if dc in self.mapped_dest_cols:
                continue

            # Command defaults check: A, B, D are populated by default
            is_cmd_default = d["code"] in MIGRATION_CONFIG["COMMAND_DEFAULTS"]

            # Find closest source column candidate
            d_candidates = [x for x in [d["name"], d["code"]] if x]
            best_source = None
            best_score = 0.0

            for sc_info in self.source_cols:
                if sc_info["col_idx"] == 1:
                    continue
                s_candidates = [x for x in [sc_info["hdr"], sc_info["ep"], sc_info["code"]] if x and not is_comment_cell(x)]
                for dc_val in d_candidates:
                    for sc_val in s_candidates:
                        score = calculate_similarity(str(dc_val), str(sc_val))
                        if score > best_score:
                            best_score = score
                            best_source = sc_info

            closest_col_let = ""
            closest_name = ""
            suggested_rename = "None"
            score_out = None

            if best_source and best_score >= threshold:
                score_out = round(best_score, 2)
                closest_col_let = best_source["col_letter"]
                closest_name = best_source["code"] or best_source["hdr"]
                suggested_rename = (
                    f'rename source column "{closest_name}" to "{d["code"] or d["name"]}"'
                )

            self.unmatched_dest.append({
                "dest_col_letter": d["col_letter"],
                "dest_col_idx": dc,
                "dest_name": d["name"],
                "dest_code": d["code"],
                "dest_inst": d["inst"],
                "is_command_default": is_cmd_default,
                "closest_col_letter": closest_col_let,
                "closest_name": closest_name,
                "suggested_rename": suggested_rename,
                "confidence_score": score_out,
            })

        self.metrics["unmatched_destination_count"] = len(self.unmatched_dest)
        self.metrics["potential_matches_with_renames"] += self.metrics["total_migrated_columns"]

    def _write_migrated_data(self) -> None:
        """Write migrated data directly into Destination.xlsx preserving rows 1-3 & formatting."""
        # 1. Create backup if requested
        if self.backup:
            backup_path = os.path.join(
                os.path.dirname(self.destination_path), "Destination_backup.xlsx"
            )
            print(f"Creating safety backup: {backup_path}")
            shutil.copyfile(self.destination_path, backup_path)

        print(f"Opening Destination.xlsx for in-place migration: {self.destination_path}")
        wb_dest = openpyxl.load_workbook(self.destination_path)
        ws_dest = wb_dest["Sheet1"]

        # Clear existing data rows (row 4 to max_row)
        max_r = max(ws_dest.max_row, 100)
        for r in range(4, max_r + 1):
            for c in range(1, ws_dest.max_column + 1):
                cell = ws_dest.cell(r, c)
                if cell.value is not None:
                    cell.value = None

        # Load source data with openpyxl
        wb_src = openpyxl.load_workbook(self.source_path, data_only=True)
        ws_src = wb_src["Sheet1"]

        # Track Assignment ID source col for Topia File Number (Col Y / Col 25)
        topia_id_sc = 24

        dest_start_row = 4
        for r_idx, src_row in enumerate(self.source_data_rows):
            dest_row = dest_start_row + r_idx

            # Populate mapped columns
            for sc, info in self.mapped_source_to_dest.items():
                dest_col_idx = info["dest"]["col_idx"]
                is_us_date = info["dest"]["is_us_date"]
                raw_val = ws_src.cell(src_row, sc).value

                if is_blank_value(raw_val):
                    clean_val = None
                elif is_us_date:
                    formatted_date, _ = format_us_date(raw_val)
                    clean_val = formatted_date
                else:
                    clean_val = raw_val

                ws_dest.cell(dest_row, dest_col_idx).value = clean_val

            # Populate command defaults (A = INSERT, B = INSERT, D = NEW)
            # Col A: 'A'
            cell_a = ws_dest.cell(dest_row, 1)
            if cell_a.value is None:
                cell_a.value = MIGRATION_CONFIG["COMMAND_DEFAULTS"]["A"]

            # Col B: 'Assignment Command'
            cell_b = ws_dest.cell(dest_row, 2)
            if cell_b.value is None:
                cell_b.value = MIGRATION_CONFIG["COMMAND_DEFAULTS"]["Assignment Command"]

            # Col D: 'Assignment Number'
            cell_d = ws_dest.cell(dest_row, 4)
            if cell_d.value is None:
                cell_d.value = MIGRATION_CONFIG["COMMAND_DEFAULTS"]["Assignment Number"]

            # Col Y: 'Topia File Number' (Col 25 / ASSIGNMENT_CF!CF_TEXT029) if empty
            cell_y = ws_dest.cell(dest_row, 25)
            if cell_y.value is None and topia_id_sc:
                cell_y.value = ws_src.cell(src_row, topia_id_sc).value

        # Ensure no extra sheets exist in Destination.xlsx
        for sname in list(wb_dest.sheetnames):
            if sname != "Sheet1":
                del wb_dest[sname]

        print(f"Saving migrated data directly into: {self.destination_path}")
        wb_dest.save(self.destination_path)
        print("Destination.xlsx updated successfully.")

    def _generate_report_workbook(self) -> None:
        """Create migration_report.xlsx with 6 beautifully formatted sheets."""
        print(f"Generating audit report workbook: {self.report_path}")
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # remove default sheet

        # Styles
        font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        fill_header = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")

        fill_green = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        font_green = Font(name="Calibri", size=10, color="276A3C", bold=True)

        fill_amber = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        font_amber = Font(name="Calibri", size=10, color="833C0C", bold=True)

        fill_red = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        font_red = Font(name="Calibri", size=10, color="C00000", bold=True)

        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )

        def style_header_row(ws, cols_count):
            ws.row_dimensions[1].height = 28
            for c_idx in range(1, cols_count + 1):
                cell = ws.cell(1, c_idx)
                cell.font = font_header
                cell.fill = fill_header
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        def autofit(ws):
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                col_let = get_column_letter(col[0].column)
                ws.column_dimensions[col_let].width = min(max(max_len + 3, 12), 65)

        # ----------------------------------------------------------------------
        # SHEET 1: Summary
        # ----------------------------------------------------------------------
        ws1 = wb.create_sheet(title="Summary")
        ws1.row_dimensions[1].height = 28
        ws1.merge_cells("A1:B1")
        ws1.cell(1, 1).value = "DATA MIGRATION RECONCILIATION SUMMARY"
        ws1.cell(1, 1).font = font_header
        ws1.cell(1, 1).fill = fill_header
        ws1.cell(1, 1).alignment = Alignment(horizontal="center", vertical="center")

        summary_rows = [
            ("Total Source Columns", self.metrics["source_columns"]),
            ("Total Destination Columns", self.metrics["destination_columns"]),
            ("Total Data Rows Migrated", self.metrics["rows_migrated"]),
            ("Exact System Code Matches", self.metrics["exact_code_matches"]),
            ("Explicit Renames (Company Levels, Group, Talent Link ID, Topia ID)", self.metrics["renamed_matches"]),
            ("Name Fallback Matches (Normalized Friendly Names)", self.metrics["name_fallback_matches"]),
            ("Duplicate Codes Disambiguated (Salary, Benefits, Title/Grade, Address)", self.metrics["duplicate_disambiguations"]),
            ("Total Migrated Columns", self.metrics["total_migrated_columns"]),
            ("Unmatched Source Columns", self.metrics["unmatched_source_count"]),
            ("Unmatched Destination Columns (Stay Blank)", self.metrics["unmatched_destination_count"]),
            ("Match count if all suggested renames are applied", self.metrics["potential_matches_with_renames"]),
        ]

        for idx, (label, val) in enumerate(summary_rows, start=2):
            c_label = ws1.cell(idx, 1)
            c_val = ws1.cell(idx, 2)
            c_label.value = label
            c_val.value = val
            c_label.font = Font(name="Calibri", size=11, bold=True)
            c_val.font = Font(name="Calibri", size=11, bold=True)
            c_label.border = thin_border
            c_val.border = thin_border
            c_val.alignment = Alignment(horizontal="right")

            if "Total Migrated" in label or "Exact" in label:
                c_label.fill = fill_green
                c_val.fill = fill_green
            elif "Renames" in label or "Fallback" in label or "Disambiguated" in label or "Match count" in label:
                c_label.fill = fill_amber
                c_val.fill = fill_amber
            elif "Unmatched" in label:
                c_label.fill = fill_red
                c_val.fill = fill_red

        ws1.column_dimensions["A"].width = 50
        ws1.column_dimensions["B"].width = 18

        # ----------------------------------------------------------------------
        # SHEET 2: Exact_Matches
        # ----------------------------------------------------------------------
        ws2 = wb.create_sheet(title="Exact_Matches")
        headers2 = [
            "Source Col Letter", "Source Col #", "Source Header (Row 4)", "Source Code (Row 2)",
            "Dest Col Letter", "Dest Col #", "Dest Friendly Name (Row 2)", "Dest Code (Row 3)",
            "Match Type", "Sample Value (Row 1)", "Sample Value (Row 2)", "Mapping Reason"
        ]
        ws2.append(headers2)
        style_header_row(ws2, len(headers2))

        for m in self.exact_matches:
            row_vals = [
                m["source_col_letter"], m["source_col_idx"], m["source_hdr"], m["source_code"],
                m["dest_col_letter"], m["dest_col_idx"], m["dest_name"], m["dest_code"],
                m["match_type"], str(m["sample_v1"] or ""), str(m["sample_v2"] or ""), m["reason"]
            ]
            ws2.append(row_vals)
            r_idx = ws2.max_row
            for c_idx in range(1, len(headers2) + 1):
                cell = ws2.cell(r_idx, c_idx)
                cell.font = Font(name="Calibri", size=10)
                cell.border = thin_border
                if c_idx == 9:  # Match Type
                    if m["match_type"] == "EXACT_CODE":
                        cell.fill = fill_green
                        cell.font = font_green
                    else:
                        cell.fill = fill_amber
                        cell.font = font_amber
        autofit(ws2)

        # ----------------------------------------------------------------------
        # SHEET 3: Not_Matched_Source
        # ----------------------------------------------------------------------
        ws3 = wb.create_sheet(title="Not_Matched_Source")
        headers3 = [
            "Source Col Letter", "Source Col #", "Source Header (Row 4)", "Source Friendly Name (Row 3)",
            "Source Code (Row 2)", "Reason Not Migrated", "Suggested Rename",
            "Target Dest Col Letter", "Target Dest Code / Name", "Confidence Score"
        ]
        ws3.append(headers3)
        style_header_row(ws3, len(headers3))

        for s in self.unmatched_source:
            row_vals = [
                s["source_col_letter"], s["source_col_idx"], s["source_hdr"], s["source_ep"],
                s["source_code"], s["reason"], s["suggested_rename"],
                s["target_col_letter"], s["target_name"], s["confidence_score"] or ""
            ]
            ws3.append(row_vals)
            r_idx = ws3.max_row
            for c_idx in range(1, len(headers3) + 1):
                cell = ws3.cell(r_idx, c_idx)
                cell.font = Font(name="Calibri", size=10)
                cell.border = thin_border
                if c_idx == 7 and s["suggested_rename"] != "None":
                    cell.fill = fill_amber
                    cell.font = font_amber
                elif c_idx == 6:
                    cell.fill = fill_red
                    cell.font = font_red
        autofit(ws3)

        # ----------------------------------------------------------------------
        # SHEET 4: Not_Matched_Destination
        # ----------------------------------------------------------------------
        ws4 = wb.create_sheet(title="Not_Matched_Destination")
        headers4 = [
            "Dest Col Letter", "Dest Col #", "Dest Friendly Name (Row 2)", "Dest Code (Row 3)",
            "Dest Instruction (Row 1)", "Command Default (A, B, D)", "Closest Source Col Letter",
            "Closest Source Header / Code", "Suggested Rename", "Confidence Score"
        ]
        ws4.append(headers4)
        style_header_row(ws4, len(headers4))

        for d in self.unmatched_dest:
            row_vals = [
                d["dest_col_letter"], d["dest_col_idx"], d["dest_name"], d["dest_code"],
                d["dest_inst"] or "", "YES (Auto-filled)" if d["is_command_default"] else "NO (Blank)",
                d["closest_col_letter"], d["closest_name"], d["suggested_rename"],
                d["confidence_score"] or ""
            ]
            ws4.append(row_vals)
            r_idx = ws4.max_row
            for c_idx in range(1, len(headers4) + 1):
                cell = ws4.cell(r_idx, c_idx)
                cell.font = Font(name="Calibri", size=10)
                cell.border = thin_border
                if c_idx == 6 and d["is_command_default"]:
                    cell.fill = fill_green
                    cell.font = font_green
                elif c_idx == 9 and d["suggested_rename"] != "None":
                    cell.fill = fill_amber
                    cell.font = font_amber
                elif c_idx == 3:
                    cell.fill = fill_red
                    cell.font = font_red
        autofit(ws4)

        # ----------------------------------------------------------------------
        # SHEET 5: Duplicates_And_Conflicts
        # ----------------------------------------------------------------------
        ws5 = wb.create_sheet(title="Duplicates_And_Conflicts")
        headers5 = [
            "System Code", "Source Col Letter", "Source Col #", "Source Header (Row 4)",
            "Resolution Status", "Target Dest Col Letter", "Target Dest Code", "Resolution Explanation"
        ]
        ws5.append(headers5)
        style_header_row(ws5, len(headers5))

        for dup in self.duplicates_log:
            sc_info = dup["sc"]
            dest_info = dup["dest"]
            row_vals = [
                dup["code"], sc_info["col_letter"], sc_info["col_idx"], sc_info["hdr"],
                dup["res"], dest_info["col_letter"] if dest_info else "None",
                dest_info["code"] if dest_info else "None", dup["notes"]
            ]
            ws5.append(row_vals)
            r_idx = ws5.max_row
            for c_idx in range(1, len(headers5) + 1):
                cell = ws5.cell(r_idx, c_idx)
                cell.font = Font(name="Calibri", size=10)
                cell.border = thin_border
                if c_idx == 5:
                    if dup["res"] == "MAPPED":
                        cell.fill = fill_green
                        cell.font = font_green
                    else:
                        cell.fill = fill_amber
                        cell.font = font_amber
        autofit(ws5)

        # ----------------------------------------------------------------------
        # SHEET 6: Data_Issues
        # ----------------------------------------------------------------------
        ws6 = wb.create_sheet(title="Data_Issues")
        headers6 = [
            "Dest Row #", "Source Row #", "Employee ID", "Employee Name", "Source Col",
            "Source Header", "Dest Col", "Dest Name", "Issue Category", "Cell Value", "Description / Recommendation"
        ]
        ws6.append(headers6)
        style_header_row(ws6, len(headers6))

        for iss in self.data_issues:
            row_vals = [
                iss["row_num"], iss["source_row"], str(iss["emp_id"] or ""), str(iss["emp_name"] or ""),
                iss["source_col"], iss["source_hdr"], iss["dest_col"], iss["dest_name"],
                iss["issue_type"], iss["value"], iss["desc"]
            ]
            ws6.append(row_vals)
            r_idx = ws6.max_row
            for c_idx in range(1, len(headers6) + 1):
                cell = ws6.cell(r_idx, c_idx)
                cell.font = Font(name="Calibri", size=10)
                cell.border = thin_border
                if c_idx == 9:
                    cell.fill = fill_red
                    cell.font = font_red
                elif c_idx == 10:
                    cell.fill = fill_amber
                    cell.font = font_amber
        autofit(ws6)

        wb.save(self.report_path)
        print(f"Audit report workbook saved: {self.report_path}")

    def print_summary(self) -> None:
        """Display console summary matching the Summary sheet."""
        print("\n" + "=" * 80)
        print("                   DATA MIGRATION RECONCILIATION SUMMARY")
        print("=" * 80)
        print(f"Source file:                         {self.source_path}")
        print(f"Destination file:                    {self.destination_path}")
        print(f"Report workbook:                     {self.report_path}")
        print("-" * 80)
        print(f"Total Source Columns:                {self.metrics['source_columns']}")
        print(f"Total Destination Columns:           {self.metrics['destination_columns']}")
        print(f"Active Data Rows Migrated:           {self.metrics['rows_migrated']}")
        print("-" * 80)
        print(f"  * Exact System Code Matches:       {self.metrics['exact_code_matches']}")
        print(f"  * Explicit Renames:                {self.metrics['renamed_matches']}")
        print(f"  * Name Fallback Matches:           {self.metrics['name_fallback_matches']}")
        print(f"  * Duplicate Code Disambiguations:  {self.metrics['duplicate_disambiguations']}")
        print(f"Total Migrated Columns:              {self.metrics['total_migrated_columns']}")
        print(f"Unmatched Source Columns:            {self.metrics['unmatched_source_count']}")
        print(f"Unmatched Destination Columns:       {self.metrics['unmatched_destination_count']}")
        print(f"Match Count If Renames Applied:      {self.metrics['potential_matches_with_renames']}")
        print(f"Data Issues Identified:              {len(self.data_issues)}")
        print("=" * 80 + "\n")


# ==============================================================================
# CLI ENTRY POINT
# ==============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Enterprise HR Data Migration: source.xlsx -> Destination.xlsx"
    )
    parser.add_argument("--source", default="source.xlsx", help="Path to source Excel file (default: source.xlsx)")
    parser.add_argument("--destination", default="Destination.xlsx", help="Path to destination Excel file (default: Destination.xlsx)")
    parser.add_argument("--report", default="migration_report.xlsx", help="Path to report Excel file (default: migration_report.xlsx)")
    parser.add_argument("--backup", action="store_true", help="Keep a safety copy of Destination.xlsx before modifying")
    parser.add_argument("--dry-run", action="store_true", help="Generate the report only, don't write to Destination.xlsx")

    args = parser.parse_args()

    try:
        migrator = DataMigrator(
            source_path=args.source,
            destination_path=args.destination,
            report_path=args.report,
            backup=args.backup,
            dry_run=args.dry_run,
        )
        migrator.run()
    except Exception as e:
        print(f"\n[ERROR] Migration failed: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
