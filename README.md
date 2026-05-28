# AIB Group Credit Risk Model

Interactive credit risk dashboard built on AIB Group's public disclosures (2020–2025).

## Live Demo
[Click here to open the dashboard](YOUR_STREAMLIT_URL_HERE)

## What It Does
- **EL Model**: Expected Loss = PD × LGD × EAD per loan segment
- **IFRS 9 Staging**: Stage 1/2/3 migration analysis and early warning signals
- **Macro Regression**: OLS model linking NPL ratios to unemployment, GDP, ECB rates
- **Stress Testing**: EBA-style Base / Adverse / Severe scenario analysis
- **PSI Analysis**: Population Stability Index to detect portfolio drift

## Data Sources
- AIB Annual Reports 2019–2024 (aib.ie/investor-relations)
- CSO Ireland (unemployment, GDP, house prices)
- ECB Data Portal (ECB rates)
- AIB Pillar 3 Reports (LGD, Stage balances)

## Tech Stack
Python · Streamlit · Plotly · Pandas · NumPy

## How to Run Locally
pip install -r requirements.txt
streamlit run app.py