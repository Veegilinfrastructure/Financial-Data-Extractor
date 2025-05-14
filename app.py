import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import re
from openpyxl import load_workbook
from transformers import pipeline

# Load FinBERT model for financial text understanding
finbert = pipeline("ner", model="ProsusAI/finbert")

def extract_text_from_pdf(pdf_file):
    """Extract text from a PDF document."""
    doc = fitz.open(pdf_file)
    text = ""
    for page in doc:
        text += page.get_text()
    return text

def extract_financial_data(text):
    """Extract key financial metrics using FinBERT and regex."""
    entities = finbert(text)

    data = {
        "Company Name": "",
        "Revenue": "",
        "EBITDA": "",
        "Industry": "",
        "Location": "",
        "Date": "",
        "Gross Profit": "",
        "Net Sales": "",
        "COGS": "",
        "Total Operating Expenses": "",
        "Operating Income": "",
        "Adjusted EBITDA": "",
    }

    # Extract key financial terms using FinBERT
    for ent in entities:
        if "ORG" in ent["entity"] and not data["Company Name"]:
            data["Company Name"] = ent["word"]
        elif "GPE" in ent["entity"] and not data["Location"]:
            data["Location"] = ent["word"]
        elif "DATE" in ent["entity"] and not data["Date"]:
            data["Date"] = ent["word"]

    # Define regex patterns for financial figures
    financial_patterns = {
        "Revenue": r"Revenue[\s:]*\$?([\d,.]+)",
        "EBITDA": r"EBITDA[\s:]*\$?([\d,.]+)",
        "Gross Profit": r"Gross Profit[\s:]*\$?([\d,.]+)",
        "Net Sales": r"Net Sales[\s:]*\$?([\d,.]+)",
        "COGS": r"COGS[\s:]*\$?([\d,.]+)",
        "Total Operating Expenses": r"Total Operating Expenses[\s:]*\$?([\d,.]+)",
        "Operating Income": r"Operating Income[\s:]*\$?([\d,.]+)",
        "Adjusted EBITDA": r"Adjusted EBITDA[\s:]*\$?([\d,.]+)"
    }

    # Extract financial values using regex
    for key, pattern in financial_patterns.items():
        match = re.search(pattern, text)
        if match:
            data[key] = match.group(1)

    return data

def save_to_excel(data):
    """Save extracted financial data to an Excel file."""
    df = pd.DataFrame([data])
    filename = "financial_data.xlsx"
    df.to_excel(filename, index=False)
    return filename

# Streamlit UI
st.title("Veegil Financial Data Extraction from PDF")

uploaded_file = st.file_uploader("Upload a PDF document", type=["pdf"])

if uploaded_file is not None:
    text = extract_text_from_pdf(uploaded_file)
    financial_data = extract_financial_data(text)
    excel_file = save_to_excel(financial_data)

    st.success("Extraction completed!")
    st.download_button("Download Excel File", excel_file, file_name="financial_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
