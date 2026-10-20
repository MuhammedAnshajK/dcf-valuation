# Nike Valuation: DCF & Comparable Companies

A valuation of NIKE, Inc. (NYSE: NKE) at 19 October 2026 using a discounted cash flow (two terminal value methods) and trading comparables against real athletic-wear peers, summarised in a football field chart. Built in Excel.

**Downloads:** [⬇ Excel model](https://github.com/MuhammedAnshajK/dcf-valuation/raw/main/nike_valuation.xlsx) · [⬇ PDF version](https://github.com/MuhammedAnshajK/dcf-valuation/raw/main/nike_valuation.pdf) · [⬇ Python script](https://github.com/MuhammedAnshajK/dcf-valuation/raw/main/build_model.py)

![Valuation summary](images/summary.png)

## What this does

The model works out what Nike is worth per share three ways and compares the result with the share price of **$37.21** (19 Oct 2026):

- **DCF, perpetuity growth (3.0%):** $37.36 per share
- **DCF, exit multiple (10.0x EBITDA):** $36.63 per share
- **Trading comps, median EV / EBITDA:** $32.90 per share; **median EV / Revenue:** $35.96

**Conclusion:** the DCF midpoint of about $37 is in line with the share price. In other words, the market is already pricing in the recovery in my forecast: revenue back to $49bn and an 11% operating margin by FY2031. A ±0.5% change in WACC or terminal growth moves the value by roughly $2 to $3 per share, so the result depends mostly on whether that margin recovery happens. The forward P/E method gives only $14.57 to $20.20 because FY2027 earnings are temporarily depressed (guidance of $1.15 to $1.35 a share), which is why Nike trades at 28x forward earnings while peers trade at about 13x.

This project builds on my [Nike Five-Year Forecast](https://github.com/MuhammedAnshajK/three-statement-model): its base case supplies the cash flows here.

## Data

| Input | Value | Source |
|---|---|---|
| Share price | $37.21 (19 Oct 2026) | stockanalysis.com |
| Shares outstanding | 1,485m | market cap $55.27bn / price |
| Beta | 1.10 (5-year) | stockanalysis.com |
| Risk-free rate | 5.09% | US 10-year Treasury, 16 Oct 2026 (YCharts) |
| Debt and cash | $7.9bn debt, $9.0bn cash and investments | Nike 10-K, 31 May 2026 |
| Forecast cash flows | FY2027 to FY2031 base case | Nike forecast model (FY2027 set to company guidance) |
| Peer multiples | adidas, Puma, Lululemon, Deckers, On, Under Armour | stockanalysis.com, 19 Oct 2026 |

The equity risk premium (5.0%) and cost of debt (5.9%) are my assumptions.

## Approach

1. **Free cash flow.** Unlevered free cash flow = EBIT x (1 - 19% tax) + D&A - capex - increase in working capital. FY2027 includes about $1bn of cash released as inventory and receivables shrink with sales.
2. **Discount rate.** Cost of equity from CAPM: 5.09% + 1.10 x 5.0% = 10.6%. After-tax cost of debt 4.8%. Market-value weights (87% equity, 13% debt) give a WACC of 9.9%.
3. **Terminal value, two ways.** Perpetuity growth at 3.0%, and an exit multiple of 10.0x FY2031 EBITDA. Each method's implied version of the other (10.3x and 2.8%) is shown as a cross-check, and they agree closely.
4. **Timing.** Cash flows are discounted from the valuation date to each 31 May year end (0.61 years to the first one).
5. **Enterprise to equity value.** Enterprise value plus net cash of $1.1bn, divided by 1,485m shares.
6. **Sensitivity.** A 5 x 5 table of value per share as WACC moves ±1% and terminal growth moves ±1% ($29.75 to $51.18).
7. **Comparables.** EV / Revenue, EV / EBITDA and forward P / E for six peers. Loss-making or negative-EBITDA multiples are excluded as not meaningful. The 25th to 75th percentile range is applied to Nike's last-twelve-month revenue and EBITDA and the midpoint of its FY2027 EPS guidance.
8. **Football field.** All ranges, plus the 52-week trading range ($31.97 to $70.44), side by side.

## Screenshots

| DCF | Sensitivity |
|---|---|
| ![DCF](images/dcf.png) | ![Sensitivity](images/sensitivity.png) |

| WACC | Comparables |
|---|---|
| ![WACC](images/wacc.png) | ![Comparables](images/comps.png) |

## How to use

1. Download `nike_valuation.xlsx` using the button at the top.
2. Change any blue or yellow cell, such as beta on the **WACC** sheet, terminal growth or the exit multiple on the **DCF** sheet, the share price on **Forecast**, or a peer multiple on **Comps**.
3. The **Summary** sheet and football field chart update automatically.

Colour key: blue = hardcoded input, black = formula, green = link to another sheet, yellow = key assumption.

## Files

- `nike_valuation.xlsx`: the model (Cover, Summary, Forecast, WACC, DCF, Sensitivity, Comps)
- `nike_valuation.pdf`: print-ready PDF of every sheet, for viewing without Excel
- `build_model.py`: Python script (openpyxl) that generates the workbook
- `images/`: screenshots

## Tools

Excel, Python (openpyxl)

*This is a learning project. It is not investment advice.*
