import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import fitz  # PyMuPDF
import io
import json
import os

# --- FUNÇÕES DE APOIO ---

def converter_pdf_para_imagens(doc_template, arquivo_pdf):
    """Converte cada página de um PDF em objetos InlineImage para o Word."""
    imagens = []
    pdf_stream = arquivo_pdf.read()
    pdf_doc = fitz.open(stream=pdf_stream, filetype="pdf")
    
    for pagina in pdf_doc:
        pix = pagina.get_pixmap(matrix=fitz.Matrix(2, 2)) # Alta qualidade
        img_byte_arr = io.BytesIO(pix.tobytes())
        # Ajusta a largura para 160mm (padrão A4 com margens)
        imagens.append(InlineImage(doc_template, img_byte_arr, width=Mm(160)))
    return imagens

def processar_imagem_simples(doc_template, arquivo_img):
    """Converte um upload de imagem em objeto InlineImage."""
    return InlineImage(doc_template, arquivo_img, width=Mm(160))

# --- INTERFACE STREAMLIT ---

st.set_page_config(page_title="Gerador de Relatórios Modelo", layout="wide")
st.title("📑 Automação de Relatórios de Prestação")

# Sidebar para seleção de contrato (conforme sua lógica de páginas diferentes)
unidade = st.sidebar.selectbox("Selecione a Unidade", ["Relatorio_Modelo"])

# Carregar Configuração (JSON que definimos antes)
# Aqui simulamos o carregamento, mas você pode usar o arquivo config.json
campos_manuais = [
    "SISTEMA_MES_REFERENCIA", "ANALISTA_TOTAL_ATENDIMENTOS", "ANALISTA_MEDICO_CLINICO",
    "ANALISTA_MEDICO_PEDIATRA", "ANALISTA_ODONTO_CLINICO", "ANALISTA_ODONTO_PED",
    "TOTAL_RAIO_X", "SISTEMA_TOTAL_DE_TRANSFERENCIA", "TOTAL_PACIENTES_CCIH",
    "OUVIDORIA_INTERNA", "OUVIDORIA_EXTERNA"
]

campos_upload = [
    "EXCEL_META_ATENDIMENTOS", "IMAGEM_PRINT_ATENDIMENTO", "IMAGEM_DOCUMENTO_RAIO_X",
    "TABELA_TRANSFERENCIA", "GRAFICO_TRANSFERENCIA", "TABELA_OBITO",
    "TABELA_TOTAL_OBITO", "TABELA_CCIH", "IMAGEM_NEP", "IMAGEM_TREINAMENTO_INTERNO",
    "IMAGEM_MELHORIAS", "GRAFICO_OUVIDORIA", "PDF_OUVIDORIA_INTERNA",
    "TABELA_QUALITATIVA_IMG", "PRINT_CLASSIFICAÇÃO"
]

# --- FORMULÁRIO ---
with st.form("form_relatorio"):
    col1, col2 = st.columns(2)
    
    contexto = {}
    
    with col1:
        st.subheader("✍️ Dados Manuais")
        for campo in campos_manuais:
            contexto[campo] = st.text_input(f"{campo.replace('_', ' ')}", key=campo)
        
        # Lógica Especial para Destinos de Transferência
        st.write("---")
        destinos_input = st.text_area("MANUAL DESTINO TRANSFERENCIA (Um por linha)")
        contexto["MANUAL_DESTINO_TRANSFERENCIA"] = " / ".join(destinos_input.split('\n'))

    with col2:
        st.subheader("📁 Upload de Arquivos")
        arquivos_upload = {}
        for campo in campos_upload:
            arquivos_upload[campo] = st.file_uploader(f"Upload para {campo}", type=['png', 'jpg', 'jpeg', 'pdf'], key=f"up_{campo}")

    enviado = st.form_submit_button("Gerar Relatório Final")

# --- PROCESSAMENTO ---
if enviado:
    try:
        # 1. Carregar Template (Certifique-se que o arquivo está na mesma pasta)
        # O arquivo deve se chamar 'template.docx' ou o nome que preferir
        template_path = "template.docx" 
        doc = DocxTemplate(template_path)
        
        # 2. Processar Cálculos Automáticos
        try:
            total = float(contexto["ANALISTA_TOTAL_ATENDIMENTOS"])
            transf = float(contexto["SISTEMA_TOTAL_DE_TRANSFERENCIA"])
            taxa = (transf / total) * 100 if total > 0 else 0
            contexto["SISTEMA_TAXA_DE_TRANSFERENCIA"] = f"{taxa:.2f}%"
        except:
            contexto["SISTEMA_TAXA_DE_TRANSFERENCIA"] = "Erro no cálculo"

        # 3. Processar Imagens e PDFs
        with st.spinner("Processando imagens e convertendo PDFs..."):
            for campo, arquivo in arquivos_upload.items():
                if arquivo:
                    if arquivo.name.lower().endswith(".pdf"):
                        contexto[campo] = converter_pdf_para_imagens(doc, arquivo)
                    else:
                        contexto[campo] = [processar_imagem_simples(doc, arquivo)]
                else:
                    contexto[campo] = [] # Lista vazia se não houver arquivo

        # 4. Renderizar e Salvar
        doc.render(contexto)
        
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        
        st.success("✅ Relatório gerado com sucesso!")
        st.download_button(
            label="📥 Baixar Relatório (.docx)",
            data=output,
            file_name=f"Relatorio_{contexto['SISTEMA_MES_REFERENCIA']}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    except Exception as e:
        st.error(f"❌ Erro ao gerar relatório: {e}")
