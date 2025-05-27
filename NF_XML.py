import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import io
from decimal import Decimal, getcontext

# Aumentar precisão
getcontext().prec = 15

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

                quantidade = Decimal(produto.find("ns:prod/ns:qCom", ns).text or '0')
                valor_unitario = Decimal(produto.find("ns:prod/ns:vUnCom", ns).text or '0')

                def get_decimal(path):
                    el = produto.find(path, ns)
                    return Decimal(el.text) if el is not None else Decimal('0')

                valor_ipi = get_decimal("ns:imposto/ns:IPI/ns:IPITrib/ns:vIPI")
                aliquota_ipi = get_decimal("ns:imposto/ns:IPI/ns:IPITrib/ns:pIPI")
                valor_icms_st = get_decimal("ns:imposto/ns:ICMS/ns:ICMS10/ns:vICMSST")
                valor_fcp_st = get_decimal("ns:imposto/ns:ICMS/ns:ICMS10/ns:vFCPST")

                valor_icms_normal = Decimal('0')
                for tag in ["ICMS00", "ICMS10", "ICMS20", "ICMS30"]:
                    el = produto.find(f"ns:imposto/ns:ICMS/ns:{tag}/ns:vICMS", ns)
                    if el is not None:
                        valor_icms_normal = Decimal(el.text)
                        break

                valor_icmsFCP_normal = Decimal('0')
                for tag in ["ICMS00", "ICMS10", "ICMS20", "ICMS30"]:
                    el = produto.find(f"ns:imposto/ns:ICMS/ns:{tag}/ns:vFCP", ns)
                    if el is not None:
                        valor_icmsFCP_normal = Decimal(el.text)
                        break

                valor_pis = get_decimal("ns:imposto/ns:PIS/ns:PISAliq/ns:vPIS")
                valor_cofins = get_decimal("ns:imposto/ns:COFINS/ns:COFINSAliq/ns:vCOFINS")

                q = quantidade if quantidade != 0 else Decimal('1')

                valor_unitario_ipi = valor_ipi / q
                valor_unitario_icms_st = valor_icms_st / q
                valor_unitario_fcp_st = valor_fcp_st / q

                valor_unitario_total = valor_unitario + valor_unitario_ipi + valor_unitario_icms_st + valor_unitario_fcp_st

                data.append({
                    "Arquivo": uploaded_file.name,
                    "Numero_Nota": numero_nota,
                    "Numero_Pedido": numero_pedido,
                    "Nome_Destinatario": nome_destinatario,
                    "Produto": nome_produto,
                    "Referencia": referencia,
                    "Quantidade_Comercial": float(quantidade),
                    "Valor_IPI": float(valor_ipi),
                    "Aliquota_IPI": float(aliquota_ipi),
                    "Valor_ICMS_ST": float(valor_icms_st),
                    "Valor_FCP_ST": float(valor_fcp_st),
                    "Valor_ICMS_Normal": float(valor_icms_normal),
                    "Valor_ICMSFCP_Normal": float(valor_icmsFCP_normal),
                    "Valor_PIS": float(valor_pis),
                    "Valor_COFINS": float(valor_cofins),
                    "Valor_Unitario": float(valor_unitario),
                    "Valor_Unitario_IPI": float(valor_unitario_ipi),
                    "Valor_Unitario_ICMS_ST": float(valor_unitario_icms_st),
                    "Valor_Unitario_ICMS_FCP_ST": float(valor_unitario_fcp_st),
                    "Valor_Unitario_Total": float(valor_unitario_total),
                    "CFOP": CFOP,
                    "Natureza": natureza,
                    "CNPJ": cnpj_destinatario.zfill(14)
                })

        except ET.ParseError:
            st.error(f"Erro ao processar o arquivo: {uploaded_file.name}")

    # DataFrame
    df = pd.DataFrame(data)

    # Garantir colunas numéricas
    num_cols = [
        'Quantidade_Comercial', 'Valor_IPI', 'Aliquota_IPI', 'Valor_ICMS_ST', 'Valor_FCP_ST',
        'Valor_ICMS_Normal', 'Valor_ICMSFCP_Normal', 'Valor_PIS', 'Valor_COFINS',
        'Valor_Unitario', 'Valor_Unitario_IPI', 'Valor_Unitario_ICMS_ST',
        'Valor_Unitario_ICMS_FCP_ST', 'Valor_Unitario_Total'
    ]
    df[num_cols] = df[num_cols].astype(float)

    # Criar resumo
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

    # Exportar para Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Detalhado', index=False)
        resumo.to_excel(writer, sheet_name='Resumo', index=False)
    output.seek(0)

    return output

# Streamlit
st.title("Processador de XMLs de Notas Fiscais — com Resumo")

uploaded_files = st.file_uploader("Faça upload dos arquivos XML", accept_multiple_files=True, type=['xml'])

if uploaded_files:
    if st.button('Processar Arquivos'):
        with st.spinner('Processando...'):
            excel_output = process_xml_files(uploaded_files)
        st.success('Processamento concluído!')

        st.download_button(
            label="Baixar Arquivo Excel",
            data=excel_output,
            file_name="valores_completos_IPI_ICMS_decimal.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
