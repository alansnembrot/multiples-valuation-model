# Multiples Valuation Model with Yahoo Finance

A Python-powered comparable companies analysis tool that fetches live financial data from Yahoo Finance and outputs a structured, formatted Excel workbook — no Bloomberg or FactSet required.

Built as a portfolio project during my transition into freelance financial analysis.

---

## What it does

Given a target company and up to 5 comparable companies (entered as Yahoo Finance tickers), the model automatically:

- Fetches key financials, valuation multiples, profitability metrics, and leverage ratios
- Builds a peer group statistics table (median, average, min, max) excluding the target
- Suggests the primary valuation multiple based on the target's sector
- Highlights the most relevant valuation variables for that sector
- Outputs an implied price range across three analyst-defined scenarios (minimum, base case, maximum)

All results are written into a clean, formatted Excel workbook with four tabs: **Input**, **Dashboard**, **Raw Data**, and **Valuation**.

---

## Output tabs

| Tab | Description |
|-----|-------------|
| **Input** | Enter target and comparable tickers |
| **Dashboard** | Side-by-side comparable analysis with peer group statistics |
| **Raw Data** | Full dataset for audit and reference |
| **Valuation** | Sector-specific variables, implied price calculator |
| **Considerations** | Methodology notes, assumptions, and known limitations |

---

## Metrics covered

**Valuation multiples** — EV/EBITDA, EV/Revenue, Price/Earnings (trailing, with forward P/E fallback)

**Size & Scale** — Market Cap, Revenue, EBITDA (in reporting currency)

**Profitability** — Gross Margin, EBITDA Margin, Return on Equity, Return on Assets

**Growth** — Revenue Growth, EBITDA Growth (2-year CAGR calculated by actual fiscal dates)

**Financial Strength** — Net Debt, Net Debt/EBITDA, Debt/Equity, EBIT/Interest Coverage

---

## Supported sectors

The model is designed for industries where EV/EBITDA and EV/Revenue are the standard valuation framework:

- Technology / Software / SaaS
- Consumer Cyclical & Consumer Defensive
- Healthcare / Pharma / Biotech
- Industrials / Manufacturing
- Energy / Oil & Gas
- Basic Materials
- Communication Services
- Utilities *(with caveats — see Considerations tab)*

**Not supported:** Financial Services (banks, insurers) and Real Estate / REITs. These sectors require different methodologies (P/BV, FFO, regulatory capital ratios) that go beyond what Yahoo Finance provides reliably. See the Considerations tab for a full explanation.

---

## Methodology notes

- **Growth rates** — 2-year CAGR calculated using actual fiscal year-end dates from Yahoo Finance financials. Empty columns and TTM periods are skipped. Handles non-December fiscal year-ends (Toyota, Volkswagen, etc.) correctly.
- **P/E ratio** — Uses trailing P/E. If trailing EPS is negative (e.g. Ford, Stellantis), falls back to forward P/E and marks the cell with `(fwd)`.
- **EV/EBITDA and EV/Revenue** — Sourced from Yahoo Finance's precalculated ratios (`enterpriseToEbitda`, `enterpriseToRevenue`) to match what appears on the Yahoo Finance company page.
- **Net Debt** — Total Debt (balance sheet) minus Cash & Cash Equivalents, pulled from the most recent balance sheet available.
- **EBIT/Interest Coverage** — Calculated manually from the income statement: EBIT / |Interest Expense|.
- **Currency** — All figures reported in each company's trading currency. No FX conversion is applied. See the Currency column.

---

## Requirements

```
Python 3.8+
yfinance
openpyxl
```

Install dependencies:
```bash
pip install yfinance openpyxl
```

---

## How to use

**1. Generate the Excel template (run once):**
```bash
python create_template.py
```
This creates `Multiples Valuation Model with Yahoo Finance.xlsx` in the same folder.

**2. Enter tickers in the Input sheet:**
Open the Excel file and enter the target company ticker in cell C6, and up to 5 comparable tickers in C9:C13.

**3. Run the model:**
```bash
python comparables_updater.py
```
Or open `comparables_updater.py` in Spyder and press F5.

**4. Open the Excel file** to see the results in Dashboard, Raw Data, and Valuation tabs.

---

## Known limitations

- Data sourced from Yahoo Finance via the unofficial `yfinance` library. Yahoo may change its data structure without notice.
- Figures are as-reported — no adjustments for one-time items or non-recurring charges.
- EBITDA growth shows N/A when the most recent year has negative EBITDA (not meaningful as a CAGR base).
- Yahoo Finance data for non-US companies is sometimes incomplete or delayed.
- With fewer than 3 comparable companies, peer group statistics have limited reliability.

Full methodology documentation is available in the **Considerations** tab of the Excel output.

---

## Project structure

```
multiples-valuation-model/
├── comparables_updater.py   # Main script — fetches data and writes Excel
├── create_template.py       # One-time setup — generates the Excel template
└── README.md
```

---

## License

This project is licensed under the **Creative Commons Attribution-NonCommercial 4.0 International (CC BY-NC 4.0)** license.

You are free to use, share, and adapt this work for non-commercial purposes, provided you give appropriate credit. Commercial use requires explicit permission from the author.

See [LICENSE](LICENSE) for full terms.

---

## Author

Built by Alan — freelance financial analyst.  
Feedback and suggestions welcome via GitHub Issues.
