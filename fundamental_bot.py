import pandas as pd
import numpy as np
from io import StringIO
import time
import os

# Selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ReportLab
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm


# Formating numbers
def format_number(value):
    try:
        if pd.isna(value) or str(value).strip() in ['-', '']:
            return "-"

        val_float = float(value)
        if val_float == 0: return "-"

        # Format with decimal comma and thousand dot
        return f"{val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(value)


# Datascrapper
def get_cvm_data(ticker):
    print(f"Extracting data for {ticker}")

    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        # Access Fundamentus
        driver.get(f"https://fundamentus.com.br/resultados_trimestrais.php?papel={ticker}&tipo=1")

        # Find the 'Exibir' (Show) link
        link_el = driver.find_element(By.PARTIAL_LINK_TEXT, "Exibir")
        driver.get(link_el.get_attribute("href"))

        # Wait for the Iframe to load
        wait = WebDriverWait(driver, 15)
        iframe = wait.until(EC.presence_of_element_located((By.ID, "iFrameFormulariosFilho")))
        driver.switch_to.frame(iframe)
        time.sleep(3)  # Safety sleep for JS execution

        # Get HTML and Parse with Pandas
        html = driver.page_source

        # header=0 assumes the first row contains the dates
        dfs = pd.read_html(StringIO(html), decimal=',', thousands='.', flavor='bs4', header=0)
        df_final = max(dfs, key=lambda d: d.size)

        # Header checker
        if isinstance(df_final.columns[0], int) or str(df_final.columns[0]) == '0':
            new_headers = df_final.iloc[0]
            df_final = df_final[1:]
            df_final.columns = new_headers

        # Drop the first column
        df_final = df_final.iloc[:, 1:].copy()

        numeric_cols = df_final.columns[1:]

        # Convert to numeric (errors become NaN)
        df_final[numeric_cols] = df_final[numeric_cols].apply(pd.to_numeric, errors='coerce')

        # Rows that have at least one valid number
        has_numbers = df_final[numeric_cols].notna().any(axis=1)

        # Rows containing specific important text (Earnings)
        is_important_text = df_final["Descrição"].astype(str).str.contains(
            "Lucro Básico|Lucro Diluído", case=False, regex=True)

        df_final = df_final[has_numbers | is_important_text]

        df_final = df_final.fillna("-")

        return df_final

    except Exception as e:
        print(f"Scraper error: {e}")
        return None
    finally:
        driver.quit()


def generate_pdf(ticker, df):
    print(f"Generating PDF for {ticker}")

    output_folder = "reports"
    os.makedirs(output_folder, exist_ok=True)

    filename = f"{ticker}_Report.pdf"
    full_path = os.path.join(output_folder, filename)

    margin_side = 5 * mm
    margin_vertical = 5 * mm

    page_width, page_height = landscape(A4)

    doc = SimpleDocTemplate(full_path, pagesize=landscape(A4),
                            leftMargin=margin_side, rightMargin=margin_side,
                            topMargin=margin_vertical, bottomMargin=margin_vertical)  # <--- MUDOU AQUI

    elements = []
    styles = getSampleStyleSheet()

    # Custom Styles
    style_num = ParagraphStyle('CN', parent=styles['Normal'], fontSize=11, leading=15, alignment=2,fontName='Helvetica-Bold')
    style_text = ParagraphStyle('CT', parent=styles['Normal'], fontSize=11, leading=15, alignment=0,fontName='Helvetica-Bold')
    style_header = ParagraphStyle('HP', parent=styles['Normal'], fontSize=11, leading=15, textColor=colors.white, alignment=1, fontName='Helvetica-Bold')

    # Title
    elements.append(Paragraph(f"Financial Report: <b>{ticker}</b> (Mil Reais)", styles['Heading3']))
    elements.append(Spacer(1, 2 * mm))  # Espaço de apenas 2mm entre título e tabela

    table_data = []

    # Headers
    headers = [Paragraph(str(col), style_header) for col in df.columns]
    table_data.append(headers)

    # Rows
    for row in df.itertuples(index=False):
        processed_row = []
        for i, item in enumerate(row):
            if i == 0:
                processed_row.append(Paragraph(str(item), style_text))
            else:
                formatted_num = format_number(item)
                processed_row.append(Paragraph(formatted_num, style_num))
        table_data.append(processed_row)

    # Widths
    available_width = page_width - (2 * margin_side)
    col_count = len(df.columns)
    data_col_count = col_count - 1

    width_desc = available_width * 0.40
    width_rest = (available_width * 0.60) / data_col_count

    col_widths = [width_desc] + [width_rest] * data_col_count

    # Table
    pdf_table = Table(table_data, colWidths=col_widths, repeatRows=1)

    # Table Style
    table_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),

        # Padding interno das células (mantive maior como você pediu antes)
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ])

    for i in range(1, len(table_data)):
        bg_color = colors.whitesmoke if i % 2 == 0 else colors.white
        table_style.add('BACKGROUND', (0, i), (-1, i), bg_color)

    pdf_table.setStyle(table_style)
    elements.append(pdf_table)

    doc.build(elements)

    print(f"Success! PDF saved at: {os.path.abspath(full_path)}")

ticker_symbol = "VALE3"
df_data = get_cvm_data(ticker_symbol)

if df_data is not None and not df_data.empty:
    generate_pdf(ticker_symbol, df_data)
else:
    print("⚠️ No data found or empty table.")