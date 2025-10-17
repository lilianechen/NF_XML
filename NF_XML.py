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

    def classificar_operacao(cfop: str, natureza: str, finNFe: str, tpNF: str) -> str:
        n = (natureza or "").lower()
        c = (cfop or "").strip()
        # finNFe: 1=normal, 2=complementar, 3=ajuste, 4=devolução
        if "remessa" in n:
            return "Remessa"
        if finNFe == "4" or "devolu" in n:
            return "Devolução"
        if "bonifica" in n:
            return "Bonificação"
        if "transfer" in n:
            return "Transferência"
        if c[:1] in {"5", "6", "7"} or tpNF == "1":
            return "Saída/Venda"
        if c[:1] in {"1", "2", "3"} or tpNF == "0":
            return "Entrada"
        return "Outras"

    for uploaded_file in uploaded_files:
        try:
            tree = ET.parse(uploaded_file)
            root = tree.getroot()

            numero_nota = (root.find(".//ns:ide/ns:nNF", ns).text
                           if root.find(".//ns:ide/ns:nNF", ns) is not None else "N/A")
            natureza = (root.find(".//ns:ide/ns:natOp", ns).text
                        if root.find(".//ns:ide/ns:natOp", ns) is not None else "N/A")
            tpNF = (root.find(".//ns:ide/ns:tpNF", ns).text
                    if root.find(".//ns:ide/ns:tpNF", ns) is not None else "")  # 0=entrada,1=saída
            finNFe = (root.find(".//ns:ide/ns:finNFe", ns).text
                      if root.find(".//ns:ide/ns:finNFe", ns) is not None else "")  # 1..4

            nome_destinatario = (root.find(".//ns:dest/ns:xNome", ns).text
                                 if root.find(".//ns:dest/ns:xNome", ns) is not None else "N/A")
            cnpj_destinatario_el = root.find(".//ns:dest/ns:CNPJ", ns)
            cpf_destinatario_el = root.find(".//ns:dest/ns:CPF", ns)
            doc_dest = (cnpj_destinatario_el.text if cnpj_destinatario_el is not None
                        else (cpf_destinatario_el.text if cpf_destinatario_el is not None else ""))
            doc_dest = (doc_dest or "").zfill(14) if doc_dest else "N/A"

            # (opcional) também captura o emitente para referência
            nome_emitente = (root.find(".//ns:emit/ns:xNome", ns).text
                             if root.find(".//ns:emit/ns:xNome", ns) is not None else "N/A")
            cnpj_emitente_el = root.find(".//ns:emit/ns:CNPJ", ns)
            cpf_emitente_el = root.find(".//ns:emit/ns:CPF", ns)
            doc_emit = (cnpj_emitente_el.text if cnpj_emitente_el is not None
                        else (cpf_emitente_el.text if cpf_emitente_el is not None else ""))
            doc_emit = (doc_emit or "").zfill(14) if doc_emit else "N/A"

            for produto in root.findall(".//ns:det", ns):
                nome_produto = produto.find("ns:prod/ns:xProd", ns).text if produto.find("ns:prod/ns:xProd", ns) is not None else "N/A"
                referencia = produto.find("ns:prod/ns:cProd", ns).text if produto.find("ns:prod/ns:cProd", ns) is not None else "N/A"
                CFOP = produto.find("ns:prod/ns:CFOP", ns).text if produto.find("ns:prod/ns:CFOP", ns) is not None else "N/A"
                numero_pedido = produto.find("ns:prod/ns:xPed", ns).text if produto.find("ns:prod/ns:xPed", ns) is not None else "N/A"

                # Quantidade / valores
                def dec_from(el):
                    return Decimal(el.text) if el is not None and el.text not in (None, "") else Decimal("0")

                quantidade = dec_from(produto.find("ns:prod/ns:qCom", ns)) or Decimal("0")
                valor_unitario = dec_from(produto.find("ns:prod/ns:vUnCom", ns))

                def get_decimal(path):
                    el = produto.find(path, ns)
                    return Decimal(el.text) if el is not None and el.text not in (None, "") else Decimal('0')

                valor_ipi = get_decimal("ns:imposto/ns:IPI/ns:IPITrib/ns:vIPI")
                aliquota_ipi = get_decimal("ns:imposto/ns:IPI/ns:IPITrib/ns:pIPI")
                valor_icms_st = (
                    get_decimal("ns:imposto/ns:ICMS/ns:ICMS10/ns:vICMSST")
                    or get_decimal("ns:imposto/ns:ICMS/ns:ICMS60/ns:vICMSST")
                )
                valor_fcp_st = (
                    get_decimal("ns:imposto/ns:ICMS/ns:ICMS10/ns:vFCPST")
                    or get_decimal("ns:imposto/ns:ICMS/ns:ICMS60/ns:vFCPST")
                )

                # ICMS normal (pega primeiro que existir)
                valor_icms_normal = Decimal('0')
                for tag in ["ICMS00", "ICMS10", "ICMS20", "ICMS30", "ICMS40", "ICMS51", "ICMS70", "ICMS90"]:
                    el = produto.find(f"ns:imposto/ns:ICMS/ns:{tag}/ns:vICMS", ns)
                    if el is not None and el.text not in (None, ""):
                        valor_icms_normal = Decimal(el.text)
                        break

                valor_icmsFCP_normal = Decimal('0')
                for tag in ["ICMS00", "ICMS10", "ICMS20", "ICMS30", "ICMS51", "ICMS70", "ICMS90"]:
                    el = produto.find(f"ns:imposto/ns:ICMS/ns:{tag}/ns:vFCP", ns)
                    if el is not None and el.text not in (None, ""):
                        valor_icmsFCP_normal = Decimal(el.text)
                        break

                valor_pis = get_decimal("ns:imposto/ns:PIS/ns:PISAliq/ns:vPIS")
                valor_cofins = get_decimal("ns:imposto/ns:COFINS/ns:COFINSAliq/ns:vCOFINS")

                q = quantidade if quantidade != 0 else Decimal('1')
                valor_unitario_ipi = valor_ipi / q
                valor_unitario_icms_st = valor_icms_st / q
                valor_unitario_fcp_st = valor_fcp_st / q
                valor_unitario_total = valor_unitario + valor_unitario_ipi + valor_unitario_icms_st + valor_unitario_fcp_st

                tipo_operacao = classificar_operacao(CFOP, natureza, finNFe, tpNF)

                data.append({
                    "Arquivo": uploaded_file.name,
                    "Numero_Nota": numero_nota,
                    "Numero_Pedido": numero_pedido,
                    "Emitente_Nome": nome_emitente,
                    "Emitente_Doc": doc_emit,
                    "Destinatario_Nome": nome_destinatario,
                    "Destinatario_Doc": doc_dest,
                    "Produto": nome_produto,
                    "Referencia": referencia,
                    "CFOP": CFOP,
                    "Natureza": natureza,
                    "tpNF": tpNF,
                    "finNFe": finNFe,
                    "Tipo_Operacao": tipo_operacao,
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
                })

        except ET.ParseError:
            st.error(f"Erro ao processar o arquivo: {uploaded_file.name}")

    # DataFrame com TUDO (sem excluir remessa!)
    df = pd.DataFrame(data)

    # Garantir colunas numéricas
    num_cols = [
        'Quantidade_Comercial', 'Valor_IPI', 'Aliquota_IPI', 'Valor_ICMS_ST', 'Valor_FCP_ST',
        'Valor_ICMS_Normal', 'Valor_ICMSFCP_Normal', 'Valor_PIS', 'Valor_COFINS',
        'Valor_Unitario', 'Valor_Unitario_IPI', 'Valor_Unitario_ICMS_ST',
        'Valor_Unitario_ICMS_FCP_ST', 'Valor_Unitario_Total'
    ]
    if not df.empty:
        df[num_cols] = df[num_cols].astype(float)

    # Resumo por nota
    resumo = (
        df.groupby('Numero_Nota', dropna=False, group_keys=False)
        .apply(lambda g: pd.Series({
            'Somatorio_Quantidade': g['Quantidade_Comercial'].sum(),
            'Somatorio_Valor_Mercadoria': (g['Quantidade_Comercial'] * g['Valor_Unitario']).sum(),
            'Total_IPI': g['Valor_IPI'].sum(),
            'Total_ICMS': g['Valor_ICMS_Normal'].sum(),
            'Total_ICMS_FCP': g['Valor_ICMSFCP_Normal'].sum(),
            'Total_PIS': g['Valor_PIS'].sum(),
            'Total_COFINS': g['Valor_COFINS'].sum(),
            'Total_ST': g['Valor_ICMS_ST'].sum() + g['Valor_FCP_ST'].sum(),
            'Somatorio_ColunaC_D': (
                (g['Quantidade_Comercial'] * g['Valor_Unitario']) +
                g['Valor_IPI'] +
                g['Valor_ICMS_Normal'] +
                g['Valor_ICMS_ST']
            ).sum(),
            # ajuda a conciliar: mostra o tipo mais frequente naquela NF
            'Tipo_Operacao_predominante': g['Tipo_Operacao'].mode().iloc[0] if not g['Tipo_Operacao'].mode().empty else ""
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
st.title("Processador de XMLs de NF-e — Todas as operações (venda, remessa, devolução, etc.)")

uploaded_files = st.file_uploader("Faça upload dos arquivos XML", accept_multiple_files=True, type=['xml'])

if uploaded_files:
    if st.button('Processar Arquivos'):
        with st.spinner('Processando...'):
            excel_output = process_xml_files(uploaded_files)
        st.success('Processamento concluído!')

        st.download_button(
            label="Baixar Arquivo Excel",
            data=excel_output,
            file_name="nfe_todas_operacoes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
