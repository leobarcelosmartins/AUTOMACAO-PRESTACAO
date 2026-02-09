import streamlit as st
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import fitz  # PyMuPDF
import io
import os
import subprocess
import tempfile
import pandas as pd
import matplotlib.pyplot as plt
from streamlit_paste_button import paste_image_button
from PIL import Image
import platform

# --- CONFIGURAÇÕES DE LAYOUT ---
st.set_page_config(page_title="Gerador de Relatórios V0.6.6", layout="wide")

# --- CUSTOM CSS PARA DASHBOARD ---
st.markdown("""
    <style>
    .main { background-color: #f0f2f5; }
    .dashboard-card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 15px;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.05);
        margin-bottom: 20px;
        border-left: 5px solid #28a745;
    }
    div.stButton > button[kind="primary"] {
        background-color: #28a745 !important;
        color: white !important;
        border: none !important;
        width: 100% !important;
        font-weight: bold !important;
        height: 3em !important;
    }
    .upload-label { font-weight: bold; color: #1f2937; margin-bottom: 8px; display: block; }
    </style>
    """, unsafe_allow_html=True)

# --- DICIONÁRIO DE DIMENSÕES ---
DIMENSOES_CAMPOS = {
    "EXCEL_META_ATENDIMENTOS": 165, "IMAGEM_PRINT_ATENDIMENTO": 165,
    "IMAGEM_DOCUMENTO_RAIO_X": 165, "TABELA_TRANSFERENCIA": 90,
    "GRAFICO_TRANSFERENCIA": 160, "TABELA_TOTAL_OBITO": 165,
    "TABELA_OBITO": 180, "TABELA_CCIH": 180, "IMAGEM_NEP": 160,
    "IMAGEM_TREINAMENTO_INTERNO": 160, "IMAGEM_MELHORIAS": 160,
    "GRAFICO_OUVIDORIA": 155, "PDF_OUVIDORIA_INTERNA": 165,
    "TABELA_QUALITATIVA_IMG": 160, "PRINT_CLASSIFICACAO": 160
}

# --- ESTADO DA SESSÃO ---
if 'dados_sessao' not in st.session_state:
    st.session_state.dados_sessao = {m: [] for m in DIMENSOES_CAMPOS.keys()}

if 'historico_capturas' not in st.session_state:
    st.session_state.historico_capturas = {m: 0 for m in DIMENSOES_CAMPOS.keys()}

def excel_para_imagem(doc_template, arquivo_excel):
    try:
        if hasattr(arquivo_excel, 'seek'): arquivo_excel.seek(0)
        df = pd.read_excel(arquivo_excel, sheet_name="TRANSFERENCIAS", usecols=[3, 4], skiprows=2, nrows=14, header=None)
        df = df.fillna('')
        def fmt(v):
            try: return str(int(float(v)))
            except: return str(v)
        if df.shape[1] > 1: df.iloc[:, 1] = df.iloc[:, 1].apply(fmt)
        fig, ax = plt.subplots(figsize=(8, 6))
        ax.axis('off')
        tabela = ax.table(cellText=df.values, loc='center', cellLoc='center', colWidths=[0.45, 0.45])
        tabela.auto_set_font_size(False)
        tabela.set_fontsize(11)
        tabela.scale(1.2, 1.8)
        for (r, c), cell in tabela.get_celld().items():
            cell.get_text().set_weight('bold')
            if r == 0: cell.set_facecolor('#D3D3D3')
        img_buf = io.BytesIO()
        plt.savefig(img_buf, format='png', bbox_inches='tight', dpi=200)
        plt.close(fig)
        img_buf.seek(0)
        return InlineImage(doc_template, img_buf, width=Mm(DIMENSOES_CAMPOS["TABELA_TRANSFERENCIA"]))
    except Exception as e:
        st.error(f"Erro Excel: {e}")
        return None

def processar_item_lista(doc_template, item, marcador):
    largura = DIMENSOES_CAMPOS.get(marcador, 165)
    try:
        if hasattr(item, 'seek'): item.seek(0)
        if isinstance(item, bytes):
            return [InlineImage(doc_template, io.BytesIO(item), width=Mm(largura))]
        ext = getattr(item, 'name', '').lower()
        if marcador == "TABELA_TRANSFERENCIA" and (ext.endswith(".xlsx") or ext.endswith(".xls")):
            res = excel_para_imagem(doc_template, item)
            return [res] if res else []
        if ext.endswith(".pdf"):
            pdf = fitz.open(stream=item.read(), filetype="pdf")
            imgs = []
            for pg in pdf:
                pix = pg.get_pixmap(matrix=fitz.Matrix(2, 2))
                imgs.append(InlineImage(doc_template, io.BytesIO(pix.tobytes()), width=Mm(largura)))
            pdf.close()
            return imgs
        return [InlineImage(doc_template, item, width=Mm(largura))]
    except Exception: return []

def converter_para_pdf(docx_path, output_dir):
    comando = 'libreoffice'
    if platform.system() == "Windows":
        caminhos_possiveis = [
            'libreoffice',
            r'C:\Program Files\LibreOffice\program\soffice.exe',
            r'C:\Program Files (x86)\LibreOffice\program\soffice.exe'
        ]
        for p in caminhos_possiveis:
            try:
                subprocess.run([p, '--version'], capture_output=True, check=True)
                comando = p
                break
            except: continue
    subprocess.run([comando, '--headless', '--convert-to', 'pdf', '--outdir', output_dir, docx_path], check=True)

# --- UI ---
st.title("Automação de Relatórios - UPA Nova Cidade")
st.caption("Versão 6.6")

t_manual, t_evidencia = st.tabs(["Dados", "Arquivos"])

with t_manual:
    st.markdown("### Preencha os campos de texto")
    
    # Reestruturação para conjuntos de 3 colunas (c1c2c3, c4c5c6, etc.)
    row1_c1, row1_c2, row1_c3 = st.columns(3)
    with row1_c1: st.text_input("Mês de Referência", key="in_mes")
    with row1_c2: st.text_input("Total de Atendimentos", key="in_total")
    with row1_c3: st.text_input("Total Raio-X", key="in_rx")
    
    row2_c4, row2_c5, row2_c6 = st.columns(3)
    with row2_c4: st.text_input("Médicos Clínicos", key="in_mc")
    with row2_c5: st.text_input("Médicos Pediatras", key="in_mp")
    with row2_c6: st.text_input("Odonto Clínico", key="in_oc")
    
    row3_c7, row3_c8, row3_c9 = st.columns(3)
    with row3_c7: st.text_input("Odonto Ped", key="in_op")
    with row3_c8: st.text_input("Pacientes CCIH", key="in_ccih")
    with row3_c9: st.text_input("Ouvidoria Interna", key="in_oi")
    
    row4_c10, row4_c11, row4_c12 = st.columns(3)
    with row4_c10: st.text_input("Ouvidoria Externa", key="in_oe")
    with row4_c11: st.number_input("Total de Transferências", step=1, key="in_tt")
    with row4_c12: st.text_input("Taxa de Transferência (%)", key="in_taxa")

with t_evidencia:
    labels = {
        "EXCEL_META_ATENDIMENTOS": "Grade de Metas", "IMAGEM_PRINT_ATENDIMENTO": "Prints Atendimento", 
        "PRINT_CLASSIFICACAO": "Classificação de Risco", "IMAGEM_DOCUMENTO_RAIO_X": "Doc. Raio-X", 
        "TABELA_TRANSFERENCIA": "Tabela Transferência (Excel)", "GRAFICO_TRANSFERENCIA": "Gráfico Transferência",
        "TABELA_TOTAL_OBITO": "Tab. Total Óbito", "TABELA_OBITO": "Tab. Óbito", 
        "TABELA_CCIH": "Tabela CCIH", "TABELA_QUALITATIVA_IMG": "Tab. Qualitativa",
        "IMAGEM_NEP": "Imagens NEP", "IMAGEM_TREINAMENTO_INTERNO": "Treinamento Interno", 
        "IMAGEM_MELHORIAS": "Melhorias", "GRAFICO_OUVIDORIA": "Gráfico Ouvidoria", 
        "PDF_OUVIDORIA_INTERNA": "Relatório Ouvidoria"
    }
    
    blocos = [
        ["EXCEL_META_ATENDIMENTOS", "IMAGEM_PRINT_ATENDIMENTO", "PRINT_CLASSIFICACAO", "IMAGEM_DOCUMENTO_RAIO_X"],
        ["TABELA_TRANSFERENCIA", "GRAFICO_TRANSFERENCIA"],
        ["TABELA_TOTAL_OBITO", "TABELA_OBITO", "TABELA_CCIH", "TABELA_QUALITATIVA_IMG"],
        ["IMAGEM_NEP", "IMAGEM_TREINAMENTO_INTERNO", "IMAGEM_MELHORIAS", "GRAFICO_OUVIDORIA", "PDF_OUVIDORIA_INTERNA"]
    ]

    for b_idx, lista_m in enumerate(blocos):
        st.markdown('<div class="dashboard-card">', unsafe_allow_html=True)
        col_esq, col_dir = st.columns(2)
        for idx, m in enumerate(lista_m):
            target = col_esq if idx % 2 == 0 else col_dir
            with target:
                st.markdown(f"<span class='upload-label'>{labels.get(m, m)}</span>", unsafe_allow_html=True)
                ca, cb = st.columns([1, 1])
                with ca:
                    key_p = f"p_{m}_{len(st.session_state.dados_sessao[m])}"
                    pasted = paste_image_button(label="Colar Print", key=key_p)
                    if pasted is not None and pasted.image_data is not None:
                        ts = getattr(pasted, 'time_now', 0)
                        if ts > st.session_state.historico_capturas[m]:
                            try:
                                img_pil = pasted.image_data
                                buf = io.BytesIO()
                                img_pil.save(buf, format="PNG")
                                b_data = buf.getvalue()
                                nome = f"Captura_{len(st.session_state.dados_sessao[m]) + 1}.png"
                                st.session_state.dados_sessao[m].append({"name": nome, "content": b_data, "type": "p"})
                                st.session_state.historico_capturas[m] = ts
                                st.rerun()
                            except: pass

                with cb:
                    f_up = st.file_uploader("Upload", type=['png', 'jpg', 'pdf', 'xlsx'], key=f"f_{m}_{b_idx}", label_visibility="collapsed")
                    if f_up:
                        if f_up.name not in [x['name'] for x in st.session_state.dados_sessao[m]]:
                            st.session_state.dados_sessao[m].append({"name": f_up.name, "content": f_up, "type": "f"})
                            st.rerun()

                if st.session_state.dados_sessao[m]:
                    for i_idx, item in enumerate(st.session_state.dados_sessao[m]):
                        with st.expander(f"{item['name']}", expanded=False):
                            if item['type'] == "p" or not item['name'].lower().endswith(('.pdf', '.xlsx')):
                                st.image(item['content'], use_container_width=True)
                            if st.button("Remover", key=f"del_{m}_{i_idx}_{b_idx}"):
                                st.session_state.dados_sessao[m].pop(i_idx)
                                st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

if st.button("FINALIZAR E GERAR RELATÓRIO", type="primary", use_container_width=True):
    mes_ref = st.session_state.get("in_mes", "").strip()
    if not mes_ref:
        st.error("Mês de Referência é obrigatório.")
    else:
        try:
            mc = int(st.session_state.get("in_mc", 0) or 0)
            mp = int(st.session_state.get("in_mp", 0) or 0)
            
            with tempfile.TemporaryDirectory() as tmp:
                docx_p = os.path.join(tmp, "relatorio.docx")
                doc = DocxTemplate("template.docx")
                with st.spinner("Processando dados e gerando arquivos..."):
                    contexto_geracao = {
                        "SISTEMA_MES_REFERENCIA": mes_ref,
                        "ANALISTA_TOTAL_ATENDIMENTOS": st.session_state.get("in_total", ""),
                        "TOTAL_RAIO_X": st.session_state.get("in_rx", ""),
                        "ANALISTA_MEDICO_CLINICO": st.session_state.get("in_mc", ""),
                        "ANALISTA_MEDICO_PEDIATRA": st.session_state.get("in_mp", ""),
                        "ANALISTA_ODONTO_CLINICO": st.session_state.get("in_oc", ""),
                        "ANALISTA_ODONTO_PED": st.session_state.get("in_op", ""),
                        "TOTAL_PACIENTES_CCIH": st.session_state.get("in_ccih", ""),
                        "OUVIDORIA_INTERNA": st.session_state.get("in_oi", ""),
                        "OUVIDORIA_EXTERNA": st.session_state.get("in_oe", ""),
                        "SISTEMA_TOTAL_DE_TRANSFERENCIA": st.session_state.get("in_tt", 0),
                        "SISTEMA_TAXA_DE_TRANSFERENCIA": st.session_state.get("in_taxa", ""),
                        "SISTEMA_TOTAL_MEDICOS": mc + mp
                    }
                    
                    for m in DIMENSOES_CAMPOS.keys():
                        imgs_doc = []
                        for item in st.session_state.dados_sessao[m]:
                            res = processar_item_lista(doc, item['content'], m)
                            if res: imgs_doc.extend(res)
                        contexto_geracao[m] = imgs_doc
                    
                    doc.render(contexto_geracao)
                    doc.save(docx_p)
                    
                    st.success("✅ Arquivos gerados com sucesso!")
                    c_down1, c_down2 = st.columns(2)
                    with c_down1:
                        with open(docx_p, "rb") as f_word:
                            st.download_button(
                                label="Baixar em WORD (.docx)",
                                data=f_word.read(),
                                file_name=f"Relatorio_{mes_ref}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                    with c_down2:
                        converter_para_pdf(docx_p, tmp)
                        pdf_final = os.path.join(tmp, "relatorio.pdf")
                        if os.path.exists(pdf_final):
                            with open(pdf_final, "rb") as f_pdf:
                                st.download_button(
                                    label="📥 Baixar em PDF",
                                    data=f_pdf.read(),
                                    file_name=f"Relatorio_{mes_ref}.pdf",
                                    mime="application/pdf",
                                    use_container_width=True
                                )
        except Exception as e: st.error(f"Erro Crítico: {e}")

st.caption("Desenvolvido por Leonardo Barcelos Martins")
