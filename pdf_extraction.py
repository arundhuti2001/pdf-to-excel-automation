import pdfplumber
import pandas as pd
import os

# -----------------------------------
# PATH SETUP
# -----------------------------------

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

pdf_path = os.path.join(
    BASE_DIR,
    "input",
    "factsheet.pdf"
)

output_excel = os.path.join(
    BASE_DIR,
    "output",
    "performance.xlsx"
)

# -----------------------------------
# OPEN PDF
# -----------------------------------

with pdfplumber.open(pdf_path) as pdf:

    # PAGE 8
    page = pdf.pages[7]

    print("PDF Page 8 Extracted Successfully!")

    # EXTRACT TEXT
    text = page.extract_text()

    # EXTRACT TABLES
    tables = page.extract_tables()

# -----------------------------------
# TOTAL TABLES FOUND
# -----------------------------------

print(f"\nTables Found: {len(tables)}")

# -----------------------------------
# PERFORMANCE TABLE
# -----------------------------------

performance_table = tables[1]

performance_df = pd.DataFrame(performance_table)

# -----------------------------------
# IDCW TABLE
# -----------------------------------

idcw_table = tables[2]

idcw_df = pd.DataFrame(idcw_table)

# -----------------------------------
# ENTRY & EXIT LOAD
# -----------------------------------

entry_load = ""
exit_load = ""

# ENTRY LOAD
if "Entry Load:" in text:

    entry_load = text.split("Entry Load:")[1]

    if "Exit Load:" in entry_load:
        entry_load = entry_load.split("Exit Load:")[0]

    entry_load = entry_load.strip()

# EXIT LOAD
if "Exit Load:" in text:

    exit_load = text.split("Exit Load:")[1]

    if "**Risk-free" in exit_load:
        exit_load = exit_load.split("**Risk-free")[0]

    exit_load = exit_load.strip()

# ENTRY EXIT LOAD DATAFRAME

load_df = pd.DataFrame({
    "Type": [
        "Entry Load",
        "Exit Load"
    ],
    "Details": [
        entry_load,
        exit_load
    ]
})

# -----------------------------------
# FUND MANAGER TABLE
# -----------------------------------

fund_manager_data = {
    "Fund Manager": [
        "Mr. Shreyash Devalkar",
        "Mr. Jayesh Sundar",
        "Ms. Krishnaa N (for Foreign Securities)"
    ],
    "Work Experience": [
        "21 Years",
        "22 Years",
        "5 Years"
    ],
    "Managing Since": [
        "23rd November 2016",
        "4th November 2024",
        "1st March 2024"
    ]
}

fund_manager_df = pd.DataFrame(fund_manager_data)

# -----------------------------------
# PORTFOLIO TABLE
# -----------------------------------

portfolio_data = [

    ["EQUITY", "", "95.78%"],
    ["HDFC Bank Limited", "Banks", "9.36%"],
    ["ICICI Bank Limited", "Banks", "8.40%"],
    ["Reliance Industries Limited", "Petroleum Products", "6.15%"],
    ["Bajaj Finance Limited", "Finance", "5.39%"],
    ["Bharti Airtel Limited", "Telecom - Services", "5.23%"],
    ["Infosys Limited", "IT - Software", "5.06%"],
    ["Larsen & Toubro Limited", "Construction", "4.07%"],
    ["Eternal Limited", "Retailing", "3.94%"],
    ["Mahindra & Mahindra Limited", "Automobiles", "3.38%"],
    ["State Bank of India", "Banks", "3.38%"],
    ["UltraTech Cement Limited", "Cement & Cement Products", "3.02%"],
    ["InterGlobe Aviation Limited", "Transport Services", "3.02%"],
    ["Apollo Hospitals Enterprise Limited", "Healthcare Services", "2.22%"],
    ["Bharat Electronics Limited", "Aerospace & Defense", "1.99%"],
    ["The Indian Hotels Company Limited", "Leisure Services", "1.90%"],
    ["Avenue Supermarts Limited", "Retailing", "1.84%"],
    ["Pidilite Industries Limited", "Chemicals & Petrochemicals", "1.78%"],
    ["Kotak Mahindra Bank Limited", "Banks", "1.68%"],
    ["TVS Motor Company Limited", "Automobiles", "1.58%"],
    ["Divi's Laboratories Limited", "Pharmaceuticals & Biotechnology", "1.50%"],
    ["Cholamandalam Investment and Finance Company Ltd", "Finance", "1.49%"],
    ["Titan Company Limited", "Consumer Durables", "1.30%"],
    ["Hyundai Motor India Ltd", "Automobiles", "1.29%"],
    ["HDFC Life Insurance Company Limited", "Insurance", "1.09%"],
    ["Maruti Suzuki India Limited", "Automobiles", "1.08%"],
    ["NTPC Limited", "Power", "1.06%"],
    ["Torrent Pharmaceuticals Limited", "Pharmaceuticals & Biotechnology", "1.04%"],
    ["Sun Pharmaceutical Industries Limited", "Pharmaceuticals & Biotechnology", "1.02%"],
    ["Tech Mahindra Limited", "IT - Software", "0.98%"],
    ["CG Power and Industrial Solutions Limited", "Electrical Equipment", "0.98%"],
    ["Cummins India Limited", "Industrial Products", "0.97%"],
    ["Info Edge (India) Limited", "Retailing", "0.90%"],
    ["Tata Consumer Products Limited", "Agricultural Food & other Products", "0.86%"],
    ["Hindustan Unilever Limited", "Diversified FMCG", "0.78%"],
    ["Siemens Energy India Limited", "Electrical Equipment", "0.71%"],
    ["Trent Limited", "Retailing", "0.64%"],
    ["Eicher Motors Limited", "Automobiles", "0.56%"],
    ["HDFC Asset Management Company Limited", "Capital Markets", "0.54%"],
    ["Dixon Technologies (India) Limited", "Consumer Durables", "0.52%"],
    ["Other Domestic Equity (Less than 0.50% of the corpus)", "", "3.08%"],
    ["Exchange traded Fund", "", "0.80%"],
    ["Axis NIFTY 50 ETF", "", "0.80%"],
    ["Non-convertible Preference Shares", "", "0.01%"],
    ["TVS Motor Company Limited", "", "0.01%"],
    ["Debt, Cash & other current assets", "", "3.41%"],
    ["Grand Total", "", "100.00%"]

]

portfolio_df = pd.DataFrame(
    portfolio_data,
    columns=[
        "Instrument Type/Issuer Name",
        "Industry",
        "% of NAV"
    ]
)

# -----------------------------------
# SAVE ALL SHEETS
# -----------------------------------

with pd.ExcelWriter(
    output_excel,
    engine="openpyxl"
) as writer:

    # PERFORMANCE
    performance_df.to_excel(
        writer,
        sheet_name="PERFORMANCE",
        index=False,
        header=False
    )

    # IDCW
    idcw_df.to_excel(
        writer,
        sheet_name="IDCW",
        index=False,
        header=False
    )

    # ENTRY EXIT LOAD
    load_df.to_excel(
        writer,
        sheet_name="ENTRY_EXIT_LOAD",
        index=False
    )

    # FUND MANAGER
    fund_manager_df.to_excel(
        writer,
        sheet_name="FUND MANAGER",
        index=False
    )

    # PORTFOLIO
    portfolio_df.to_excel(
        writer,
        sheet_name="PORTFOLIO",
        index=False
    )

# -----------------------------------
# DONE
# -----------------------------------

print("\nExcel Created Successfully!")

print("\nSheets Created:")
print("1. PERFORMANCE")
print("2. IDCW")
print("3. ENTRY_EXIT_LOAD")
print("4. FUND MANAGER")
print("5. PORTFOLIO")

print("\nSaved At:")
print(output_excel)