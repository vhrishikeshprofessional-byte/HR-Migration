# HR Data Migration Tool (`source.xlsx` -> `Destination.xlsx`)

Migrates employee assignment data from `source.xlsx` directly into the existing `Destination.xlsx` template layout while generating a detailed 6-sheet reconciliation report (`migration_report.xlsx`).

---

## Quick Start

### 1. Dry Run (Preview & generate report without touching Destination)
```bash
python migrate.py --dry-run
```

### 2. Live Run (Migrates data directly into Destination.xlsx)
```bash
python migrate.py
```

### 3. Live Run with Safety Backup (Recommended)
```bash
python migrate.py --backup
```
*(Creates `Destination_backup.xlsx` before writing).*

---

## CLI Options

| Argument | Default | Description |
| :--- | :--- | :--- |
| `--source` | `source.xlsx` | Path to source workbook |
| `--destination` | `Destination.xlsx` | Path to target template workbook |
| `--report` | `migration_report.xlsx` | Path to generated audit report |
| `--backup` | `False` | Saves `Destination_backup.xlsx` before modifying |
| `--dry-run` | `False` | Audits and generates report only (read-only mode) |

---

## How It Works (Mapping Priority)

The script maps columns in strict priority order to prevent collisions:

1. **Exact System Code Match**: Matches Source Row 2 code with Destination Row 3 code.
2. **Explicit Renames**: Mapped via `MIGRATION_CONFIG["RENAMES"]` at the top of `migrate.py` (e.g. `FROM_COMPANY_LEVEL1..4`, `GROUP`, `SYSTEM_NUMBER`, and legacy `Assignment ID` -> `Topia Id`).
3. **Duplicate Code Disambiguation**: Uses Source Row 4 headers to distinguish duplicate codes:
   - Base Pay Currency vs. Base Pay Amount
   - Host Job Title vs. Host Job Level / Salary Grade
   - Benefits Currency vs. Benefits Amount
   - Work Home Address vs. other location types
   - Primary Phone vs. Mobile Phone
4. **Name Fallback**: If Source Row 2 is blank or a comment, matches normalized friendly names (Source Row 3/4 == Destination Row 2).
5. **Fuzzy Suggestions**: Suggests closest column renames in the report for any unmapped columns (never auto-migrated).

### Data Handling
- **Dates**: Formatted as `MM/DD/YYYY` for columns marked `"US Date Format"` in Destination Row 1.
- **Data Types**: IDs, numbers, and strings with leading zeros are preserved as-is.
- **Blanks**: Empty strings and literal `"None"` are treated as empty cells.
- **Command Columns**: Destination columns `A` (`INSERT`), `B` (`INSERT`), and `D` (`NEW`) are auto-filled if empty.
- **Template Safety**: Destination rows 1–3, styling, and dropdown validations are never modified.

---

## Outputs

- **`Destination.xlsx`**: Updated in-place with migrated data rows starting at Row 4.
- **`migration_report.xlsx`**: 6-sheet audit workbook:
  - `Summary`: Overall match statistics and potential match count.
  - `Exact_Matches`: Itemized list of all migrated columns and sample values.
  - `Not_Matched_Source`: Unmapped source columns with rename suggestions.
  - `Not_Matched_Destination`: Unmapped destination columns with closest candidates.
  - `Duplicates_And_Conflicts`: Duplicate source codes and their exact resolution.
  - `Data_Issues`: Flagged data anomalies (e.g., corrupted text, country code mismatches, dropdown casing).
