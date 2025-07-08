import streamlit as st
import xlwings as xw

EXCEL_PATH = "calculator.xlsx"  # File must be in the same folder

def read_excel_inputs():
    wb = xw.Book(EXCEL_PATH)
    sheet = wb.sheets[0]
    values = [sheet.range(f'B{i}').value for i in range(1, 26)]
    wb.close()
    return values

def write_inputs_and_calculate(values):
    wb = xw.Book(EXCEL_PATH)
    sheet = wb.sheets[0]
    for i, val in enumerate(values, start=1):
        sheet.range(f'B{i}').value = val
    result = sheet.range('B26').value
    wb.close()
    return result

st.set_page_config(page_title="Casting Calculator")

st.title("Casting Price Calculator")

labels = [
    "Casting weight", "Cavity Nr.", "Scrap ratio", "Total core weight", "Yield",
    "Grinding time", "Filter", "EXO Feeder", "C", "Si", "Mn", "P", "S", "Cr", "Cu", "Mg",
    "Length", "Width", "Tickness", "Moulds/h", "Energy Cost", "Monthly Salary",
    "Profit margin", "Annual Amortization", "General cost"
]

default_values = read_excel_inputs()

with st.form("calc_form"):
    inputs = []
    for label, default in zip(labels, default_values):
        val = st.number_input(label, value=float(default), format="%.4f")
        inputs.append(val)
    submitted = st.form_submit_button("Calculate")

if submitted:
    result = write_inputs_and_calculate(inputs)
    st.success(f"Final price per piece: €{result:.4f}")

