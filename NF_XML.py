import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import io
from decimal import Decimal, getcontext
import zipfile
from io import BytesIO

# Aumentar precisão
getcontext().prec = 15

def classificar_operacao(cfop: str, natureza: str, finNFe: str, tpNF: str) -> str:
    n = (natureza or "").lower()
    c = (cfop or "").strip()
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

def process_single_xml(xml_content, filename):
    """Processa um único arquivo XML e retorna os dados"""
    ns = {"ns": "http://www.portalfiscal.inf.br/nfe"}
    data = []
    
    try:
        tree = ET.parse(BytesIO(xml_content))
        root = tree.getroot()

        numero_nota = (root.find(".//ns:ide/ns:nNF", ns).text
                       if root.find(".//ns:ide/ns:nNF", ns) is not None else "N/A")
        natureza = (root.find(".//ns:ide/ns:natOp", ns).text
                    if root.find(".//ns:ide/ns:natOp", ns) is not None else "N/A")
        tpNF = (root.find(".//ns:ide/ns:tpNF", ns).text
                if root.find(".//ns:ide/ns:tpNF", ns) is not None else "")
        finNFe = (root.find(".//ns:ide/ns:finNFe", ns).text
                  if root.find(".//ns:ide/ns:finNFe", ns) is not None else "")

        nome_destinatario = (root.find(".//ns:dest/ns:xNome", ns).text
                             if root.find(".//ns:dest/ns:xNome", ns) is not None else "N/A")
        cnpj_destinatario_el = root.find(".//ns:dest/ns:CNPJ", ns)
        cpf_destinatario_el = root.find(".//ns:dest/ns:CPF", ns)
        doc_dest = (cnpj_destinatario_el.text if cnpj_destinatario_el is not None
                    else (cpf_destinatario_el.text if cpf_destinatario_el is not None else ""))
        doc_dest = (doc_dest or "").zfill(14) if doc_dest else "N/A"

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
                "Arquivo": filename,
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

    except Exception as e:
        st.warning(f"Erro ao processar {filename}: {str(e)}")
    
    return data

def process_xml_files(uploaded_files, progress_bar, status_text):
    """Processa todos os arquivos XML com barra de progresso"""
    all_data = []
    total_files = len(uploaded_files)
    processed = 0
    errors = 0
    
    for idx, uploaded_file in enumerate(uploaded_files):
        try:
            # Ler o conteúdo do arquivo
            xml_content = uploaded_file.read()
            
            # Processar o XML
            data = process_single_xml(xml_content, uploaded_file.name)
            all_data.extend(data)
            processed += 1
            
        except Exception as e:
            st.warning(f"Erro ao processar {uploaded_file.name}: {str(e)}")
            errors += 1
        
        # Atualizar progresso
        progress = (idx + 1) / total_files
        progress_bar.progress(progress)
        status_text.text(f"Processando: {idx + 1}/{total_files} arquivos ({processed} OK, {errors} erros)")
    
    # Criar DataFrame
    df = pd.DataFrame(all_data)
    
    # FILTRAR: Remover notas de remessa
    if not df.empty:
        total_antes = len(df)
        df = df[df['Tipo_Operacao'] != 'Remessa'].copy()
        total_depois = len(df)
        remessas_removidas = total_antes - total_depois
        if remessas_removidas > 0:
            st.info(f"🗑️ {remessas_removidas} linha(s) de notas de Remessa foram removidas do processamento")
    
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
            'Tipo_Operacao_predominante': g['Tipo_Operacao'].mode().iloc[0] if not g['Tipo_Operacao'].mode().empty else ""
        }), include_groups=False)
        .reset_index()
    )

    # Exportar para Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Detalhado', index=False)
        resumo.to_excel(writer, sheet_name='Resumo', index=False)
    output.seek(0)

    return output, processed, errors

def extract_zip(zip_file):
    """Extrai XMLs de um arquivo ZIP"""
    xml_files = []
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        for file_info in zip_ref.filelist:
            if file_info.filename.lower().endswith('.xml'):
                xml_content = zip_ref.read(file_info.filename)
                # Criar objeto tipo arquivo
                xml_file = BytesIO(xml_content)
                xml_file.name = file_info.filename
                xml_files.append(xml_file)
    return xml_files

# Configuração da página
st.set_page_config(
    page_title="Processador NF-e",
    page_icon="📄",
    layout="wide"
)

st.title("📄 Processador de XMLs de NF-e")
st.markdown("### Processamento de vendas, devoluções, bonificações, etc. (exceto Remessa)")

# Aumentar limite de upload
st.markdown("""
**💡 Dicas para processar muitos arquivos:**
- Para mais de 200 XMLs, use um arquivo ZIP
- Limite recomendado: até 5000 arquivos por processamento
- Arquivos ZIP podem ter até 200MB
- ⚠️ **Notas de Remessa são automaticamente excluídas do processamento**
""")

# Tabs para diferentes métodos de upload
tab1, tab2 = st.tabs(["📁 Upload de XMLs", "📦 Upload de ZIP"])

with tab1:
    uploaded_files = st.file_uploader(
        "Faça upload dos arquivos XML (máximo 200 arquivos)",
        accept_multiple_files=True,
        type=['xml'],
        key="xml_uploader"
    )
    
    if uploaded_files:
        st.info(f"✅ {len(uploaded_files)} arquivo(s) carregado(s)")
        
        if st.button('🚀 Processar Arquivos XML', type="primary"):
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            excel_output, processed, errors = process_xml_files(uploaded_files, progress_bar, status_text)
            
            st.success(f'✅ Processamento concluído! {processed} arquivos processados, {errors} erros')
            
            st.download_button(
                label="⬇️ Baixar Arquivo Excel",
                data=excel_output,
                file_name="nfe_todas_operacoes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

with tab2:
    zip_file = st.file_uploader(
        "Faça upload de um arquivo ZIP contendo XMLs",
        type=['zip'],
        key="zip_uploader"
    )
    
    if zip_file:
        st.info("📦 Arquivo ZIP carregado. Clique para extrair e processar.")
        
        if st.button('🚀 Extrair e Processar ZIP', type="primary"):
            with st.spinner('Extraindo XMLs do ZIP...'):
                xml_files = extract_zip(zip_file)
            
            st.success(f'✅ {len(xml_files)} arquivos XML extraídos do ZIP')
            
            if xml_files:
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                excel_output, processed, errors = process_xml_files(xml_files, progress_bar, status_text)
                
                st.success(f'✅ Processamento concluído! {processed} arquivos processados, {errors} erros')
                
                st.download_button(
                    label="⬇️ Baixar Arquivo Excel",
                    data=excel_output,
                    file_name="nfe_todas_operacoes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )

# Informações adicionais
with st.expander("ℹ️ Sobre o processador"):
    st.markdown("""
    Este processador extrai informações de XMLs de NF-e e gera um relatório em Excel com:
    
    **Aba Detalhado:**
    - Informações de cada produto em cada nota
    - Dados fiscais completos (IPI, ICMS, PIS, COFINS, etc.)
    - Classificação automática do tipo de operação
    
    **Aba Resumo:**
    - Totalizadores por nota fiscal
    - Valores consolidados de impostos
    - Tipo de operação predominante
    
    **Tipos de operação identificados:**
    - ~~Remessa~~ (automaticamente excluída)
    - Devolução
    - Bonificação
    - Transferência
    - Saída/Venda
    - Entrada
    - Outras
    """)
