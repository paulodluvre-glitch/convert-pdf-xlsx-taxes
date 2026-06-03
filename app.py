import re
import io
from datetime import datetime

import pandas as pd
import pdfplumber
import streamlit as st


# ============================================================
# CONFIGURAÇÃO DO APP
# ============================================================

st.set_page_config(
    page_title="Conversor de Guias de Impostos",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Conversor de Guias de Impostos em Excel")
st.write(
    "Suba uma ou várias guias em PDF. O app extrai os dados principais, "
    "organiza os impostos em tabela e gera uma planilha Excel útil para conferência."
)


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def br_money_to_float(value):
    """
    Converte valores no formato brasileiro:
    '8.611,54' -> 8611.54
    """
    if pd.isna(value):
        return 0.0

    value = str(value).strip()

    if not value:
        return 0.0

    value = value.replace(".", "").replace(",", ".")

    try:
        return float(value)
    except ValueError:
        return 0.0


def float_to_br_money(value):
    """
    Converte float para texto no padrão brasileiro.
    """
    try:
        return f"{value:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return value


def extract_text_from_pdf(uploaded_file):
    """
    Extrai texto de todas as páginas do PDF usando pdfplumber.
    Funciona bem para PDFs textuais da Receita Federal.
    """
    full_text = ""

    with pdfplumber.open(uploaded_file) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            full_text += f"\n\n--- PAGE {page_number} ---\n{text}"

    return full_text


def extract_header_data(text, file_name):
    """
    Extrai dados gerais da guia:
    CNPJ, razão social, número do documento, pagar até, valor total.
    """
    header = {
        "arquivo": file_name,
        "cnpj": "",
        "razao_social": "",
        "numero_documento": "",
        "pagar_ate": "",
        "valor_total_documento": 0.0,
        "periodo_apuracao_cabecalho": "",
        "observacoes": ""
    }

    # CNPJ
    cnpj_match = re.search(r"\b\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\b", text)
    if cnpj_match:
        header["cnpj"] = cnpj_match.group(0)

    # Número do documento
    doc_match = re.search(r"\b\d{2}\.\d{2}\.\d{5}\.\d{7}-\d\b", text)
    if doc_match:
        header["numero_documento"] = doc_match.group(0)

    # Data "Pagar até"
    pagar_ate_match = re.search(r"Pagar(?: este documento)? até\s*(\d{2}/\d{2}/\d{4})", text, flags=re.I)
    if not pagar_ate_match:
        pagar_ate_match = re.search(r"Pagar até:\s*(\d{2}/\d{2}/\d{4})", text, flags=re.I)

    if pagar_ate_match:
        header["pagar_ate"] = pagar_ate_match.group(1)

    # Valor total
    valor_total_match = re.search(
        r"Valor Total do Documento\s*([\d\.\,]+)",
        text,
        flags=re.I
    )

    if not valor_total_match:
        valor_total_match = re.search(
            r"Valor:\s*([\d\.\,]+)",
            text,
            flags=re.I
        )

    if valor_total_match:
        header["valor_total_documento"] = br_money_to_float(valor_total_match.group(1))

    # Razão social
    # Pega a linha onde aparece CNPJ + nome.
    if header["cnpj"]:
        pattern_razao = re.escape(header["cnpj"]) + r"\s+(.+)"
        razao_match = re.search(pattern_razao, text)

        if razao_match:
            razao = razao_match.group(1).strip()

            # Remove pedaços indesejados caso venham grudados
            razao = re.split(
                r"Período de Apuração|Data de Vencimento|Número do Documento|Pagar",
                razao
            )[0].strip()

            header["razao_social"] = razao

    # Período de apuração do cabeçalho, quando vier como "Diversos"
    periodo_match = re.search(r"Período de Apuração\s+(.+?)\s+Data de Vencimento", text, flags=re.I | re.S)
    if periodo_match:
        periodo = periodo_match.group(1).replace("\n", " ").strip()
        header["periodo_apuracao_cabecalho"] = periodo

    # Observações
    obs_match = re.search(r"Observações\s+(.+?)\s+Valor Total do Documento", text, flags=re.I | re.S)
    if obs_match:
        obs = obs_match.group(1).replace("\n", " ").strip()
        header["observacoes"] = obs

    return header


def parse_tax_items(text, header):
    """
    Extrai as linhas da composição da guia.

    Estrutura esperada:
    Código Denominação Principal Multa Juros Total
    0561 IRRF - RENDIMENTO DO TRABALHO ASSALARIADO 39,76 7,95 1,74 49,45
    07 IRRF - ...
    PA 10/2025 Vencimento 19/11/2025
    """

    rows = []

    lines = []
    for raw_line in text.splitlines():
        line = raw_line.strip()

        if not line:
            continue

        # Remove marcações internas
        if line.startswith("--- PAGE"):
            continue

        # Ignora rodapé/cabeçalho repetitivo
        ignored_starts = (
            "Documento de Arrecadação",
            "de Receitas Federais",
            "Código Denominação",
            "Composição do Documento",
            "SENDA",
            "AUTENTICAÇÃO MECÂNICA",
            "Pague com o PIX",
            "Receita Federal"
        )

        if line.startswith(ignored_starts):
            continue

        lines.append(line)

    current_item = None

    # Linha principal do imposto:
    # 0561 IRRF - RENDIMENTO DO TRABALHO ASSALARIADO 39,76 7,95 1,74 49,45
    main_line_pattern = re.compile(
        r"^(\d{4})\s+(.+?)\s+"
        r"([\d\.]+,\d{2})\s+"
        r"([\d\.]+,\d{2})\s+"
        r"([\d\.]+,\d{2})\s+"
        r"([\d\.]+,\d{2})$"
    )

    # Linha do PA:
    # PA 10/2025 Vencimento 19/11/2025
    pa_pattern = re.compile(
        r"PA\s+([0-9]{2}/[0-9]{4}|[0-9]{4})\s+Vencimento\s+(\d{2}/\d{2}/\d{4})",
        flags=re.I
    )

    for line in lines:
        main_match = main_line_pattern.match(line)

        if main_match:
            # Se tinha um item pendente sem PA, salva mesmo assim
            if current_item:
                rows.append(current_item)

            codigo = main_match.group(1)
            denominacao = main_match.group(2).strip()

            current_item = {
                "arquivo": header["arquivo"],
                "cnpj": header["cnpj"],
                "razao_social": header["razao_social"],
                "numero_documento": header["numero_documento"],
                "pagar_ate": header["pagar_ate"],
                "valor_total_documento": header["valor_total_documento"],
                "codigo": codigo,
                "denominacao": denominacao,
                "detalhamento": "",
                "periodo_apuracao": "",
                "vencimento_original": "",
                "principal": br_money_to_float(main_match.group(3)),
                "multa": br_money_to_float(main_match.group(4)),
                "juros": br_money_to_float(main_match.group(5)),
                "total": br_money_to_float(main_match.group(6)),
            }

            continue

        pa_match = pa_pattern.search(line)

        if pa_match and current_item:
            current_item["periodo_apuracao"] = pa_match.group(1)
            current_item["vencimento_original"] = pa_match.group(2)

            rows.append(current_item)
            current_item = None

            continue

        # Linha intermediária do detalhamento
        # Exemplo: 01 CP SEGURADOS - EMPREGADOS/AVULSO
        if current_item:
            if not line.startswith(("Totais", "8587")):
                if current_item["detalhamento"]:
                    current_item["detalhamento"] += " | " + line
                else:
                    current_item["detalhamento"] = line

    # Salva o último item caso tenha ficado pendente
    if current_item:
        rows.append(current_item)

    return rows


def process_pdf(uploaded_file):
    """
    Processa um PDF e retorna:
    - header
    - dataframe de lançamentos
    - texto bruto extraído
    """
    file_name = uploaded_file.name

    text = extract_text_from_pdf(uploaded_file)
    header = extract_header_data(text, file_name)
    rows = parse_tax_items(text, header)

    df = pd.DataFrame(rows)

    return header, df, text


def create_excel_file(df_all, headers_df):
    """
    Gera Excel em memória com:
    - Lançamentos
    - Resumo por código
    - Resumo por PA
    - Resumo por guia
    - Validação
    """

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Aba 1: lançamentos detalhados
        df_all.to_excel(writer, index=False, sheet_name="Lancamentos")

        # Aba 2: resumo por código
        resumo_codigo = (
            df_all
            .groupby(["codigo", "denominacao"], dropna=False)
            .agg(
                principal=("principal", "sum"),
                multa=("multa", "sum"),
                juros=("juros", "sum"),
                total=("total", "sum"),
                qtd_linhas=("codigo", "count")
            )
            .reset_index()
            .sort_values(["codigo", "denominacao"])
        )
        resumo_codigo.to_excel(writer, index=False, sheet_name="Resumo por Codigo")

        # Aba 3: resumo por período de apuração
        resumo_pa = (
            df_all
            .groupby(["periodo_apuracao"], dropna=False)
            .agg(
                principal=("principal", "sum"),
                multa=("multa", "sum"),
                juros=("juros", "sum"),
                total=("total", "sum"),
                qtd_linhas=("periodo_apuracao", "count")
            )
            .reset_index()
            .sort_values(["periodo_apuracao"])
        )
        resumo_pa.to_excel(writer, index=False, sheet_name="Resumo por PA")

        # Aba 4: resumo por guia/documento
        resumo_guia = (
            df_all
            .groupby(
                [
                    "arquivo",
                    "cnpj",
                    "razao_social",
                    "numero_documento",
                    "pagar_ate",
                    "valor_total_documento"
                ],
                dropna=False
            )
            .agg(
                soma_principal=("principal", "sum"),
                soma_multa=("multa", "sum"),
                soma_juros=("juros", "sum"),
                soma_total_lancamentos=("total", "sum"),
                qtd_linhas=("total", "count")
            )
            .reset_index()
        )

        resumo_guia["diferenca_documento_vs_lancamentos"] = (
            resumo_guia["valor_total_documento"] - resumo_guia["soma_total_lancamentos"]
        )

        resumo_guia.to_excel(writer, index=False, sheet_name="Resumo por Guia")

        # Aba 5: cabeçalhos extraídos
        headers_df.to_excel(writer, index=False, sheet_name="Cabecalhos")

        # Aba 6: validação
        validacao = resumo_guia[
            [
                "arquivo",
                "numero_documento",
                "valor_total_documento",
                "soma_total_lancamentos",
                "diferenca_documento_vs_lancamentos",
                "qtd_linhas"
            ]
        ].copy()

        validacao["status"] = validacao["diferenca_documento_vs_lancamentos"].apply(
            lambda x: "OK" if abs(x) <= 0.01 else "DIVERGENTE"
        )

        validacao.to_excel(writer, index=False, sheet_name="Validacao")

        # Formatação básica
        workbook = writer.book

        money_columns = {
            "Lancamentos": ["valor_total_documento", "principal", "multa", "juros", "total"],
            "Resumo por Codigo": ["principal", "multa", "juros", "total"],
            "Resumo por PA": ["principal", "multa", "juros", "total"],
            "Resumo por Guia": [
                "valor_total_documento",
                "soma_principal",
                "soma_multa",
                "soma_juros",
                "soma_total_lancamentos",
                "diferenca_documento_vs_lancamentos"
            ],
            "Validacao": [
                "valor_total_documento",
                "soma_total_lancamentos",
                "diferenca_documento_vs_lancamentos"
            ]
        }

        for sheet_name in workbook.sheetnames:
            ws = workbook[sheet_name]

            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            # Cabeçalho
            for cell in ws[1]:
                cell.font = cell.font.copy(bold=True)
                cell.alignment = cell.alignment.copy(horizontal="center")

            # Largura automática simples
            for col in ws.columns:
                max_length = 0
                col_letter = col[0].column_letter

                for cell in col:
                    value = cell.value
                    if value is not None:
                        max_length = max(max_length, len(str(value)))

                ws.column_dimensions[col_letter].width = min(max_length + 2, 45)

            # Formato monetário
            if sheet_name in money_columns:
                headers = [cell.value for cell in ws[1]]

                for money_col in money_columns[sheet_name]:
                    if money_col in headers:
                        col_index = headers.index(money_col) + 1

                        for row in range(2, ws.max_row + 1):
                            ws.cell(row=row, column=col_index).number_format = '#,##0.00'

    output.seek(0)
    return output


# ============================================================
# INTERFACE STREAMLIT
# ============================================================

uploaded_files = st.file_uploader(
    "Suba as guias em PDF",
    type=["pdf"],
    accept_multiple_files=True
)

if uploaded_files:
    all_rows = []
    headers = []
    extraction_logs = []

    with st.spinner("Processando PDFs..."):
        for uploaded_file in uploaded_files:
            try:
                header, df, text = process_pdf(uploaded_file)

                headers.append(header)

                if not df.empty:
                    all_rows.append(df)

                extraction_logs.append({
                    "arquivo": uploaded_file.name,
                    "status": "OK",
                    "linhas_extraidas": len(df),
                    "valor_total_documento": header.get("valor_total_documento", 0.0)
                })

            except Exception as e:
                extraction_logs.append({
                    "arquivo": uploaded_file.name,
                    "status": f"ERRO: {str(e)}",
                    "linhas_extraidas": 0,
                    "valor_total_documento": 0.0
                })

    logs_df = pd.DataFrame(extraction_logs)

    st.subheader("Resultado do processamento")
    st.dataframe(logs_df, use_container_width=True)

    if all_rows:
        df_all = pd.concat(all_rows, ignore_index=True)
        headers_df = pd.DataFrame(headers)

        # Ordenação útil
        sort_cols = [
            "razao_social",
            "numero_documento",
            "periodo_apuracao",
            "codigo",
            "denominacao"
        ]

        existing_sort_cols = [col for col in sort_cols if col in df_all.columns]
        df_all = df_all.sort_values(existing_sort_cols).reset_index(drop=True)

        st.subheader("Prévia dos lançamentos extraídos")
        st.dataframe(df_all, use_container_width=True)

        st.subheader("Resumo por código")
        resumo_codigo = (
            df_all
            .groupby(["codigo", "denominacao"], dropna=False)
            .agg(
                principal=("principal", "sum"),
                multa=("multa", "sum"),
                juros=("juros", "sum"),
                total=("total", "sum"),
                qtd_linhas=("codigo", "count")
            )
            .reset_index()
            .sort_values(["codigo", "denominacao"])
        )

        st.dataframe(resumo_codigo, use_container_width=True)

        excel_file = create_excel_file(df_all, headers_df)

        file_name = f"guias_impostos_convertidas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        st.download_button(
            label="📥 Baixar Excel organizado",
            data=excel_file,
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.success("Planilha gerada com sucesso.")

    else:
        st.error(
            "Nenhuma linha de imposto foi extraída. "
            "Provavelmente o PDF é imagem/scaneado ou tem layout muito diferente."
        )
else:
    st.info("Suba um ou mais PDFs para começar.")