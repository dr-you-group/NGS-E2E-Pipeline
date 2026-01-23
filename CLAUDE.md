# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

NGS-E2E-Pipeline is a FastAPI web application that processes NGS (Next-Generation Sequencing) Excel reports and generates structured HTML reports. The application parses TruSight Oncology 500 (TSO500) Excel files containing genomic variant data and stores them in SQLite for report generation and retrieval.

## Running the Application

**Start the development server:**
```bash
python app.py
```
The server runs on `http://0.0.0.0:1234` with hot reload enabled.

**Python version:** Python 3.13 (uses venv in `venv/` directory)

**Install dependencies:** No requirements.txt exists. Key dependencies include:
- fastapi
- uvicorn
- pandas
- openpyxl
- jinja2

## Architecture

### Core Components

**app.py** - Main FastAPI application with the following key endpoints:
- `GET /` - Main upload page
- `POST /api/upload-excel` - Processes uploaded Excel files
- `GET /report/{specimen_id}` - Displays generated report
- `GET /api/search` - Searches reports by specimen_id
- `GET /api/reports` - Lists all reports

**utils.py** - `NGS_EXCEL2DB` class that parses TSO500 Excel files with multiple sheets:
- `clinical_information` - Patient/specimen metadata
- `SNV`, `CNV`, `Fusion`, `Splice`, `LR_BRCA` - Variant data sheets
- `NGS_QC`, `IO` - Quality control and biomarkers
- Extracts data into structured dictionaries for report generation

### Data Flow

1. User uploads Excel file via web interface
2. File temporarily saved to `tmp/` directory
3. `NGS_EXCEL2DB` parses all sheets and extracts structured data
4. Report data JSON stored in SQLite (`ngs_reports.db`) with specimen_id as key
5. Optionally saved as JSON file in `json/` directory
6. Temporary Excel file cleaned up after processing

### Storage

- **SQLite Database:** `ngs_reports.db` - Single table `reports` with columns: id, specimen_id (unique), report_data (JSON blob), created_at
- **JSON Files:** `json/{specimen_id}.json` - Backup storage of report data
- **Temporary Files:** `tmp/` - Uploaded Excel files (auto-cleaned after processing)

### Panel Types

The application supports two panel types determined by specimen type:
- **GE** (General) - TruSight Oncology 500 DNA/RNA panel
- **SA** (Special) - TSO500 + RNA Fusion Panel

Panel type affects reagent specifications and gene content displayed in reports.

### Report Data Structure

Reports include:
- Clinical information (patient demographics, specimen details)
- Variants by type: SNV/Indels, CNV, Fusion, Splice, LR_BRCA (large rearrangements in BRCA1/2)
- Each variant type split into VCS (clinically significant) and VUS (unknown significance)
- Biomarkers: TMB, MSI
- QC metrics, filter history, analysis program details
- User information (tested by, signed by, analyzed by)

### Split Info Feature

Large variant tables (SNV, CNV) use `process_table_data_with_split_info()` to indicate pagination:
- `split_at` field marks where to split data between first and subsequent pages
- SNV/CNV: max 8-10 rows on first page
- Fusion/Splice/LR_BRCA: no splitting (typically fewer entries)

## Important Implementation Notes

**File Locking:** The application explicitly calls `parser.close()` after processing to release Excel file locks. Use `safe_remove_file()` for reliable temporary file cleanup with retries.

**Specimen ID:** Extracted from `병리번호` field in clinical_information sheet. Used as primary key for all database operations.

**Database Operations:** Always use `INSERT OR REPLACE` to allow report updates with same specimen_id.

**Template System:** Uses Jinja2 templates in `templates/` directory. Specification and gene content HTML files served as static endpoints.

**Error Handling:** Use `safe_remove_file()` with retries for temporary file cleanup. Always wrap parser operations in try/except and call `parser.close()` in finally blocks to prevent file locks.

## API Endpoints

Additional endpoints beyond the main ones:
- `GET /api/specification/{panel_type}` - Returns specification HTML for GE or SA panels
- `GET /api/gene-content/{content_type}` - Returns gene content HTML (valid types: GE_Gene_Content_DRNA, SA_Gene_Content_DNA, SA_Gene_Content_RNA)
- `POST /generate-report` - Alternative endpoint to generate report from specimen_id form submission
