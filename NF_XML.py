import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import io

# Função para processar arquivos XML
def process_xml_files(uploaded_files):
    data = []
    ns = {"ns": "http://www.portalfiscal.inf.br/nfe"}

    for uploaded_file in uploaded_files:
        try:
            tree = ET.parse(uploaded_file)
            root = tree.getroot()

            numero_nota = root.find(".//ns:ide/ns:nNF", ns).text if root.find(".//ns:ide/ns:nNF", ns) is not None else "N/A"
            natureza = root.find(".//ns:ide/ns:natOp", ns).text if root.find(".//ns:ide/ns:natOp", ns) is not None else "N/A"
            nome_destinatario = root.find(".//ns:dest/ns:xNome", ns).text if root.find(".//ns:dest/ns:xNome", ns) is not None else "N/A"
            cnpj_destinatario = root.find(".//ns:dest/ns:CNPJ", ns).text if root.find(".//ns:dest/ns:CNPJ", ns) is not None else "N/A"

            for produto in root.findall(".//ns:det", ns):
                nome_produto = produto.find("ns:prod/ns:xProd", ns).text if produto.find("ns:prod/ns:xProd", ns) is not None else "N/A"
                referencia = produto.find("ns:prod/ns:cProd", ns).text if produto.find("ns:prod/ns:cProd", ns) is not None else "N/A"
                CFOP = produto.find("ns:prod/ns:CFOP", ns).text if produto.find("ns:prod/ns:CFOP", ns) is not None else "N/A"
                numero_pedido = produto.find("ns:prod/ns:xPed", ns).text if produto.find("ns:prod/ns:xPed", ns) is not None else "N/A"
                quantidade_comercial = produto.find("ns:prod/ns:qCom", ns).text if produto.find("ns:prod/ns:qCom", ns) is not None else "0"
                valor_ipi = produto.find("ns:imposto/ns:IPI/ns:IPITrib/ns:vIPI", ns)
                valor_ipi = valor_ipi.text if valor_ipi is not None else "0"
                aliquota_ipi = produto.find("ns:imposto/ns:IPI/ns:IPITrib/ns:pIPI", ns)
                aliquota_ipi = aliquota_ipi.text if aliquota_ipi is not None else "0"
                valor_icms_st = produto.find("ns:imposto/ns:ICMS/ns:ICMS10/ns:vICMSST", ns)
                valor_icms_st = valor_icms_st.text if valor_icms_st is not None else "0"
                valor_fcp_st = produto.find("ns:imposto/ns:ICMS/ns:ICMS10/ns:vFCPST", ns)
                valor_fcp_st = valor_fcp_st.text if valor_fcp_st is not None else "0"

                icms_normal = None
                for tag in ["ICMS00", "ICMS10", "ICMS20", "ICMS30"]:
                    icms_normal = produto.find(f"ns:imposto/ns:ICMS/ns:{tag}/ns:vICMS", ns)
                    if icms_normal is not None:
                        break
                valor_icms_normal = icms_normal.text if icms_normal is not None else "0"

                icms_fcp = None
                for tag in ["ICMS00", "ICMS10", "ICMS20", "ICMS30"]:
                    icms_fcp = produto.find(f"ns:imposto/ns:ICMS/ns:{tag}/ns:vFCP", ns)
                    if icms_fcp is not None:
                        break
                valor_icmsFCP_normal = icms_fcp.text if icms_fcp is not None else "0"

                valor_pis = produto.find("ns:imposto/ns:PIS/ns:PISAliq/ns:vPIS", ns)
                valor_pis = valor_pis.text if valor_pis is not None else "0"
                valor_cofins = produto.find("ns:imposto/ns:COFINS/ns:COFINSAliq/ns:vCOFINS", ns)
                valor_cofins = valor_cofins.text if valor_cofins is not None else "0"
                valor_unitario = produto.find("ns:prod/ns:vUnCom", ns).text if produto.find("ns:prod/ns:vUnCom", ns) is not None else "0"

                data.append({
                    "Arquivo": uploaded_file.name,
                    "Numero_Nota": numero_nota,
                    "Numero_Pedido": numero_pedido,
                    "Nome_Destinatario": nome_destinatario,
                    "Produto": nome_produto,
                    "Referencia": referencia,
                    "Quantidade_Comercial": quantidade_comercial,
                    "Valor_IPI": valor_ipi,
                    "Aliquota_IPI": aliquota_ipi,
                    "Valor_ICMS_ST": valor_icms_st,
                    "Valor_FCP_ST": valor_fcp_st,
                    "Valor_ICMS_Normal": valor_icms_normal,
                    "Valor_ICMSFCP_Normal": valor_icmsFCP_normal,
                    "Valor_PIS": valor_pis,
                    "Valor_COFINS": valor_cofins,
                    "Valor_Unitario": valor_unitario,
                    "CFOP": CFOP,
                    "Natureza": natureza,
                    "CNPJ": cnpj_destinatario.zfill(14)
                })
        except ET.ParseError:
            st.error(f"Erro ao processar o arquivo: {uploaded_file.name}")

    df = pd.DataFrame(data, dtype=str)
    df = df[~df['Natureza'].str.startswith('Remessa', na=False)]

    # Conversão numérica
    def convert_to_float(column):
        return column.str.replace(',', '.', regex=False).astype(float)

    num_cols = ['Quantidade_Comercial', 'Valor_IPI', 'Valor_ICMS_Normal', 'Valor_PIS', 'Valor_COFINS',
                'Valor_Unitario', 'Aliquota_IPI', 'Valor_ICMS_ST', 'Valor_FCP_ST', 'Valor_ICMSFCP_Normal']
    for col in num_cols:
        df[col] = convert_to_float(df[col])

    # Cálculos adicionais
    df['Valor_Unitario_Total'] = (
        df['Valor_IPI'] / df['Quantidade_Comercial'] +
        df['Valor_ICMS_Normal'] / df['Quantidade_Comercial'] +
        df['Valor_Unitario'] +
        df['Valor_ICMS_ST'] / df['Quantidade_Comercial']
    )

    df['Valor_Unitario_ICMS_ST'] = df['Valor_ICMS_ST'] / df['Quantidade_Comercial']
    df['Valor_Unitario_ICMS_FCP_ST'] = df['Valor_FCP_ST'] / df['Quantidade_Comercial']

    # Resumo corrigido
    resumo = (
        df.groupby('Numero_Nota', group_keys=False)
        .apply(lambda group: pd.Series({
            'Somatorio_Quantidade': group['Quantidade_Comercial'].sum(),
            'Somatorio_Valor_Mercadoria': (group['Quantidade_Comercial'] * group['Valor_Unitario']).sum(),
            'Total_IPI': group['Valor_IPI'].sum(),
            'Total_ICMS': group['Valor_ICMS_Normal'].sum(),
            'Total_ICMS_FCP': group['Valor_ICMSFCP_Normal'].sum(),
            'Total_PIS': group['Valor_PIS'].sum(),
            'Total_COFINS': group['Valor_COFINS'].sum(),
            'Total_ST': group['Valor_ICMS_ST'].sum() + group['Valor_FCP_ST'].sum(),
            'Somatorio_ColunaC_D': (
                (group['Quantidade_Comercial'] * group['Valor_Unitario']) +
                group['Valor_IPI'] +
                group['Valor_ICMS_Normal'] +
                group['Valor_ICMS_ST']
            ).sum()
        }))
        .reset_index()
    )

    # Criação do arquivo Excel em memória
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Detalhado', index=False)
        resumo.to_excel(writer, sheet_name='Resumo', index=False)
    output.seek(0)

    return output

# Interface Streamlit
st.title("Processador de XMLs de Notas Fiscais")

uploaded_files = st.file_uploader("Faça upload dos arquivos XML", accept_multiple_files=True, type=['xml'])

if uploaded_files:
    if st.button('Processar Arquivos'):
        with st.spinner('Processando...'):
            excel_output = process_xml_files(uploaded_files)
        st.success('Processamento concluído!')

        st.download_button(
            label="Baixar Arquivo Excel",
            data=excel_output,
            file_name="valores_completos_IPI_ICMS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
