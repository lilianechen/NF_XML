import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import io
from decimal import Decimal, getcontext

# Aumentar precisão
getcontext().prec = 15

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
                
                # Conversão com Decimal
                quantidade = Decimal(produto.find("ns:prod/ns:qCom", ns).text)
                valor_ipi = Decimal(produto.find("ns:imposto/ns:IPI/ns:IPITrib/ns:vIPI", ns).text or '0')
                aliquota_ipi = Decimal(produto.find("ns:imposto/ns:IPI/ns:IPITrib/ns:pIPI", ns).text or '0')
                valor_icms_st = Decimal(produto.find("ns:imposto/ns:ICMS/ns:ICMS10/ns:vICMSST", ns).text or '0')
                valor_fcp_st = Decimal(produto.find("ns:imposto/ns:ICMS/ns:ICMS10/ns:vFCPST", ns).text or '0')
                valor_unitario = Decimal(produto.find("ns:prod/ns:vUnCom", ns).text or '0')

                icms_normal = None
                for tag in ["ICMS00", "ICMS10", "ICMS20", "ICMS30"]:
                    icms_normal_element = produto.find(f"ns:imposto/ns:ICMS/ns:{tag}/ns:vICMS", ns)
                    if icms_normal_element is not None:
                        icms_normal = Decimal(icms_normal_element.text)
                        break
                valor_icms_normal = icms_normal if icms_normal is not None else Decimal('0')

                icms_fcp = None
                for tag in ["ICMS00", "ICMS10", "ICMS20", "ICMS30"]:
                    icms_fcp_element = produto.find(f"ns:imposto/ns:ICMS/ns:{tag}/ns:vFCP", ns)
                    if icms_fcp_element is not None:
                        icms_fcp = Decimal(icms_fcp_element.text)
                        break
                valor_icmsFCP_normal = icms_fcp if icms_fcp is not None else Decimal('0')

                valor_pis = Decimal(produto.find("ns:imposto/ns:PIS/ns:PISAliq/ns:vPIS", ns).text or '0')
                valor_cofins = Decimal(produto.find("ns:imposto/ns:COFINS/ns:COFINSAliq/ns:vCOFINS", ns).text or '0')

                # Evita divisão por zero
                q = quantidade if quantidade != 0 else Decimal('1')

                # Cálculos unitários
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

    # DataFrame com números como número
    df = pd.DataFrame(data)

    # Exportar para Excel como número
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Detalhado', index=False)
    output.seek(0)

    return output

# Interface Streamlit
st.title("Processador de XMLs de Notas Fiscais — Números como número")

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
