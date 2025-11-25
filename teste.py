import pandas as pd
import numpy as np
from io import StringIO
import time
import os
import sys

# Selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ReportLab
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm


def format_number(value):
    try:
        if pd.isna(value) or str(value).strip() in ['-', '']:
            return "-"

        val_float = float(value)
        if val_float == 0:
            return "-"
        # Format with decimal comma and thousand dot
        return f"{val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(value)

def get_cvm_data_all(ticker):

    print(f"Getting report data for {ticker}.")

    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--window-size=1920,1080")

    # Results
    results = {
        "Ativo": None,
        "Passivo": None,
        "DRE": None
    }

    # Searching in Select
    reports_config = [
        ("Ativo", ["Balanço", "Ativo"], "Ativo|Circulante|Caixa|Imobilizado"),
        ("Passivo", ["Balanço", "Passivo"], "Passivo|Patrimônio|Obrigações|Provisões"),
        ("DRE", ["Demonstração", "Resultado"], "Receita|Lucro|Resultado|Despesas|Ebitda")
    ]

    driver = None

    try:
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
        wait = WebDriverWait(driver, 20)

        # Fundamentus website
        url = f"https://fundamentus.com.br/resultados_trimestrais.php?papel={ticker}&tipo=1"
        driver.get(url)

        link_el = wait.until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Exibir")))
        driver.get(link_el.get_attribute("href"))
        time.sleep(3)

        # 2. Loop para extrair cada relatório
        for report_name, select_keywords, filter_regex in reports_config:
            print(f"\n--- Processando: {report_name} ---")

            # A. SELECIONAR NO MENU (Contexto Principal)
            # Importante: Garantir que estamos no contexto default antes de buscar o select
            driver.switch_to.default_content()

            try:
                # Tenta achar o select
                try:
                    select_element = wait.until(EC.element_to_be_clickable((By.ID, "cmbQuadro")))
                except:
                    select_element = wait.until(EC.element_to_be_clickable((By.TAG_NAME, "select")))

                select_obj = Select(select_element)

                # Achar a opção correta
                target_text = None
                for option in select_obj.options:
                    text = option.text.strip()
                    # Verifica se TODAS as palavras chaves estão no texto da opção
                    if all(k in text for k in select_keywords):
                        target_text = text
                        break

                if target_text:
                    # Só muda se não for a que já está selecionada (otimização)
                    if select_obj.first_selected_option.text != target_text:
                        print(f"Selecionando menu: {target_text}")
                        select_obj.select_by_visible_text(target_text)
                        print("Aguardando atualização da tabela...")
                        time.sleep(5)  # Tempo para o AJAX da CVM rodar
                    else:
                        print(f"Menu já selecionado: {target_text}")
                else:
                    print(f"⚠️ Opção para {report_name} não encontrada no menu.")
                    continue

            except Exception as e:
                print(f"Erro ao manipular menu para {report_name}: {e}")
                continue

            # B. LER DADOS (Contexto do iFrame)
            try:
                # Entra no iframe
                iframe = wait.until(EC.presence_of_element_located((By.ID, "iFrameFormulariosFilho")))
                driver.switch_to.frame(iframe)

                # Pega HTML
                html = driver.page_source

                # Parse
                dfs = pd.read_html(StringIO(html), decimal=',', thousands='.', flavor='bs4', header=0)

                if dfs:
                    df_final = max(dfs, key=lambda d: d.size)

                    # Limpeza de Cabeçalho
                    if isinstance(df_final.columns[0], int) or str(df_final.columns[0]) == '0':
                        new_headers = df_final.iloc[0]
                        df_final = df_final[1:]
                        df_final.columns = new_headers

                    df_final = df_final.iloc[:, 1:].copy()  # Remove coluna checkbox
                    numeric_cols = df_final.columns[1:]
                    df_final[numeric_cols] = df_final[numeric_cols].apply(pd.to_numeric, errors='coerce')

                    # Filtros
                    has_numbers = df_final[numeric_cols].notna().any(axis=1)
                    is_summary = df_final.iloc[:, 0].astype(str).str.contains(filter_regex, case=False, regex=True)

                    df_final = df_final[has_numbers | is_summary].fillna("-")

                    # Salva no dicionário
                    results[report_name] = df_final
                    print(f"✅ {report_name} extraído com sucesso ({len(df_final)} linhas).")
                else:
                    print(f"❌ Nenhuma tabela encontrada para {report_name}.")

            except Exception as e:
                print(f"Erro ao ler tabela de {report_name}: {e}")
                # Importante: se der erro no iframe, tenta voltar pro default pra não quebrar o próximo loop
                driver.switch_to.default_content()

    except Exception as e:
        print(f"Erro geral no WebDriver: {e}")
    finally:
        if driver:
            driver.quit()

    return results


def generate_consolidated_pdf(ticker, data_dict):
    print(f"\nGerando PDF consolidado para {ticker}...")

    output_folder = "reports"
    os.makedirs(output_folder, exist_ok=True)

    filename = f"{ticker}_Report.pdf"
    full_path = os.path.join(output_folder, filename)

    margin_side = 5 * mm
    margin_vertical = 5 * mm

    doc = SimpleDocTemplate(full_path, pagesize=landscape(A4),
                            leftMargin=margin_side, rightMargin=margin_side,
                            topMargin=margin_vertical, bottomMargin=margin_vertical)

    elements = []
    styles = getSampleStyleSheet()

    # Custom Styles
    style_num = ParagraphStyle('CN', parent=styles['Normal'], fontSize=7, leading=9, alignment=2, fontName='Helvetica')
    style_text = ParagraphStyle('CT', parent=styles['Normal'], fontSize=7, leading=9, alignment=0,
                                fontName='Helvetica-Bold')
    style_header = ParagraphStyle('HP', parent=styles['Normal'], fontSize=8, leading=10, textColor=colors.white,
                                  alignment=1, fontName='Helvetica-Bold')
    style_title = styles['Heading2']
    style_title.alignment = 1  # Center

    # Ordem de impressão
    report_order = ["Ativo", "Passivo", "DRE"]

    first_page = True

    for report_name in report_order:
        df = data_dict.get(report_name)

        if df is None or df.empty:
            continue

        if not first_page:
            elements.append(PageBreak())
        first_page = False

        # Título da Seção
        title_map = {
            "Ativo": "Balanço Patrimonial - Ativo",
            "Passivo": "Balanço Patrimonial - Passivo",
            "DRE": "Demonstração do Resultado do Exercício"
        }

        elements.append(
            Paragraph(f"{title_map.get(report_name, report_name)}: <b>{ticker}</b> (Mil Reais)", style_title))
        elements.append(Spacer(1, 3 * mm))

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
                    processed_row.append(Paragraph(format_number(item), style_num))
            table_data.append(processed_row)

        # Widths Calculation
        page_width, _ = landscape(A4)
        available_width = page_width - (2 * margin_side)
        col_count = len(df.columns)
        data_col_count = col_count - 1

        if data_col_count > 0:
            width_desc = available_width * 0.25
            width_rest = (available_width * 0.75) / data_col_count
            col_widths = [width_desc] + [width_rest] * data_col_count
        else:
            col_widths = [available_width]

        # Table Object
        pdf_table = Table(table_data, colWidths=col_widths, repeatRows=1)

        # Styling
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 7),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ])

        # Zebra striping
        for i in range(1, len(table_data)):
            bg_color = colors.whitesmoke if i % 2 == 0 else colors.white
            table_style.add('BACKGROUND', (0, i), (-1, i), bg_color)

        pdf_table.setStyle(table_style)
        elements.append(pdf_table)

    try:
        doc.build(elements)
        print(f"Sucesso! PDF consolidado salvo em: {os.path.abspath(full_path)}")
    except Exception as e:
        print(f"Erro ao gerar PDF: {e}")


# Main Execution Block
if __name__ == "__main__":
    ticker_symbol = "VALE3"

    # 1. Extrair todos os dados
    all_data = get_cvm_data_all(ticker_symbol)

    # 2. Verificar se temos dados para gerar o PDF
    if any(df is not None and not df.empty for df in all_data.values()):
        generate_consolidated_pdf(ticker_symbol, all_data)
    else:
        print("⚠️ Nenhum dado foi extraído. Verifique os erros acima.")