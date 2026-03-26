# Lending Institution Profiler

**Objective:** Map the lending profiles of top Indian banks, NBFCs, and DFIs — who lends what, to whom, and where they're headed.

## Institutions Covered (Top 15)
| # | Name | Type |
|---|------|------|
| 1 | State Bank of India | PSU Bank |
| 2 | HDFC Bank | Private Bank |
| 3 | Bank of Baroda | PSU Bank |
| 4 | Punjab National Bank | PSU Bank |
| 5 | Canara Bank | PSU Bank |
| 6 | Union Bank of India | PSU Bank |
| 7 | Indian Bank | PSU Bank |
| 8 | Central Bank of India | PSU Bank |
| 9 | Bank of India | PSU Bank |
| 10 | Bank of Maharashtra | PSU Bank |
| 11 | IDFC First Bank | Private Bank |
| 12 | LIC Housing Finance Ltd | NBFC |
| 13 | SIDBI | DFI |
| 14 | Export-Import Bank of India | DFI |
| 15 | National Housing Bank | DFI |

## Lending Instruments Tracked
- INR Term Loans
- Working Capital Lines / CC/OD
- External Commercial Borrowings (ECB)
- Direct Assignment / DA
- Pass-Through Certificates (PTC) / Securitisation
- Trade Finance (LC/BG)
- G-Sec / SDL Purchases
- Corporate Loans (CPS/CCPS)
- Infrastructure Financing
- SME/MSME Lending

## Project Structure
```
lending-institution-profiler/
├── README.md
├── requirements.txt
├── data/
│   └── raw/           # Scraped raw data
│   └── processed/     # Cleaned outputs
├── src/
│   ├── __init__.py
│   ├── scraper.py     # Web scraping (press releases, annual reports)
│   ├── annual_report.py # Annual report data extraction
│   ├── processor.py   # Data cleaning & structuring
│   └── excel_output.py # Excel workbook generator
├── outputs/
│   └── [dated Excel files]
└── notebooks/
    └── analysis.ipynb
```

## Usage
```bash
pip install -r requirements.txt
python src/scraper.py          # Scrape public data
python src/processor.py       # Clean and structure
python src/excel_output.py    # Generate Excel workbook
```

## Data Sources
- Institution websites / investor presentations
- RBI数据库 (RBI database)
- Prime MBA / Bloomberg
- Money Control / Economic Times
- CIBIL / Experian India
- Annual Reports (MCA21)

## Goals
1. ✅ Current lending profile (instrument × institution matrix)
2. ⬜ Sector focus & strategic direction (press releases, transcripts)
3. ⬜ Sector exposure concentration
4. ⬜ Relationship mapping (existing shared borrowers)
5. ⬜ Pitch deck generation per institution
