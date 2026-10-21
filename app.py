import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Pt, Cm
import io
import os
import copy
from pptx.opc.constants import RELATIONSHIP_TYPE as RT


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def duplicate_slide(presentation, index):
    source = presentation.slides[index]
    dest = presentation.slides.add_slide(source.slide_layout)

    for shp in list(dest.shapes):
        shp._element.getparent().remove(shp._element)

    src_spTree = source.shapes._spTree
    dst_spTree = dest.shapes._spTree

    for child in list(dst_spTree):
        tag = child.tag.split('}')[-1]
        if tag not in ('nvGrpSpPr', 'grpSpPr'):
            dst_spTree.remove(child)

    for child in list(src_spTree):
        tag = child.tag.split('}')[-1]
        if tag in ('nvGrpSpPr', 'grpSpPr'):
            continue
        new_el = copy.deepcopy(child)
        dst_spTree.append(new_el)

    rels_map = {}
    for rel in source.part.rels.values():
        if rel.reltype == RT.SLIDE_LAYOUT:
            continue
        if 'notesSlide' in rel.reltype:
            continue
        if rel.is_external:
            new_rid = dest.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            new_rid = dest.part.relate_to(rel.target_part, rel.reltype)
        rels_map[rel.rId] = new_rid

    for elem in dst_spTree.iter():
        for attr in list(elem.attrib.keys()):
            if attr.endswith('}embed') or attr.endswith('}link') or attr.endswith('}id'):
                old_rid = elem.get(attr)
                if old_rid and old_rid in rels_map:
                    elem.set(attr, rels_map[old_rid])

    return dest


def delete_slide(presentation, index):
    xml_slides = presentation.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[index])


def substitute_text_in_slide(slide, substituicoes, campos_para_limpar=None, apagar_shape=False, font_size=None):
    campos_para_limpar = campos_para_limpar or []
    chaves_limpar_upper = [c.upper() for c in campos_para_limpar]

    def substituir_no_run(run, subs):
        texto = run.text
        alterou = False
        for chave, valor in subs.items():
            if chave in texto or chave.capitalize() in texto or chave.lower() in texto:
                texto = texto.replace(chave, str(valor))
                texto = texto.replace(chave.capitalize(), str(valor))
                texto = texto.replace(chave.lower(), str(valor))
                alterou = True
        
        run.text = texto
        
        # Se alterou o texto e o usuário definiu um tamanho de fonte, aplica
        if alterou and font_size is not None:
            run.font.size = Pt(font_size)

    def limpar_no_run(run, chaves):
        texto = run.text
        for chave in chaves:
            for variacao in (chave, chave.capitalize(), chave.lower()):
                if variacao in texto:
                    texto = texto.replace(variacao, "")
        run.text = texto

    def processar_text_frame(tf):
        for paragraph in tf.paragraphs:
            for run in paragraph.runs:
                substituir_no_run(run, substituicoes)
                limpar_no_run(run, chaves_limpar_upper)

    def processar_shape(shape):
        if shape.has_text_frame and chaves_limpar_upper:
            texto_shape = shape.text_frame.text.upper()
            contem_limpar = any(chave in texto_shape for chave in chaves_limpar_upper)

            if contem_limpar and apagar_shape:
                texto_restante = texto_shape
                for chave in chaves_limpar_upper:
                    texto_restante = texto_restante.replace(chave, "")
                if not texto_restante.strip(" \t\n\r-:|·•"):
                    shape._element.getparent().remove(shape._element)
                    return

        if shape.has_text_frame:
            processar_text_frame(shape.text_frame)

        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    processar_text_frame(cell.text_frame)

        if shape.shape_type == 6:  # GROUP
            for sub in list(shape.shapes):
                processar_shape(sub)

    for shape in list(slide.shapes):
        processar_shape(shape)


# ============================================================
# INTERFACE STREAMLIT
# ============================================================

st.set_page_config(page_title="Gerador de Prismas", page_icon="🔷", layout="centered")

st.title("🔷 Gerador de Prismas")
st.caption("Versão Web 1.3 - Criado por João Carlos")

st.markdown("---")

# PASSO 1: Upload do Excel
st.subheader("1️⃣ Arquivo Excel")
excel_file = st.file_uploader("Envie o arquivo Excel", type=["xlsx", "xls"])

if not excel_file:
    st.info("Aguardando o arquivo Excel para começar...")
    st.stop()

try:
    xls = pd.ExcelFile(excel_file)
    sheet_names = xls.sheet_names
    st.success(f"✅ Excel carregado: {len(sheet_names)} aba(s) encontrada(s)")
except Exception as e:
    st.error(f"Erro ao ler Excel: {e}")
    st.stop()

# PASSO 2: Escolher aba e Mapeamento
st.subheader("2️⃣ Selecionar Aba")
selected_sheet = st.selectbox("Aba do Excel", sheet_names)

df = pd.read_excel(excel_file, sheet_name=selected_sheet)
colunas = [''] + df.columns.tolist()

st.markdown("**Mapeamento de Colunas**")
col1, col2, col3 = st.columns(3)
with col1:
    nome_col = st.selectbox("Nome", colunas, index=0)
with col2:
    cargo_col = st.selectbox("Cargo (opcional)", colunas, index=0)
with col3:
    empresa_col = st.selectbox("Empresa", colunas, index=0)

# --- NOVIDADE: Configuração de Fonte ---
st.markdown("**Configurações de Texto e Layout**")
col_conf1, col_conf2 = st.columns(2)
with col_conf1:
    apagar_shape_vazio = st.checkbox("Apagar caixa se não usada", value=False, help="Remove a caixa de texto inteira se ela só continha o campo vazio (ex: apenas a palavra CARGO).")
with col_conf2:
    usar_tamanho_custom = st.checkbox("Alterar tamanho da fonte")
    
font_size = None
if usar_tamanho_custom:
    font_size = st.slider("Tamanho da fonte (Pt)", min_value=8, max_value=100, value=24, step=2, help="Aplicado apenas aos campos substituídos (Nome, Cargo, Empresa).")

st.markdown("---")

# PASSO 3: Formato e Template
st.subheader("3️⃣ Formato e Template")

# NOVIDADE: Seleção de Tamanho
formato_prisma = st.radio(
    "Selecione o formato do prisma:",
    ["21x14,5 cm", "30x10,5 cm"],
    horizontal=True
)

modo_template = st.radio(
    "Como deseja escolher o template?",
    ["Usar template padrão", "Enviar meu template"],
    horizontal=True,
)

TEMPLATES_PADRAO = {
    "SESI": "sesi.pptx",
    "SENAI": "senai.pptx",
    "CNI": "cni.pptx",
    "IEL": "iel.pptx",
    "Sistema Indústria": "sistema_industria.pptx",
}

pasta_tamanho = "21x14.5" if formato_prisma == "21x14,5 cm" else "30x10.5"

template_file = None
template_path = None
nome_template = None
forcar_redimensionamento = False

if modo_template == "Usar template padrão":
    nome_template = st.selectbox("Escolha o modelo", list(TEMPLATES_PADRAO.keys()))
    # O caminho agora inclui a pasta de formato!
    template_path = os.path.join("templates", pasta_tamanho, TEMPLATES_PADRAO[nome_template])

    if os.path.exists(template_path):
        st.success(f"✅ Template **{nome_template}** ({formato_prisma}) carregado")
    else:
        st.error(f"❌ Arquivo `{template_path}` não encontrado. Verifique a pasta.")
else:
    template_file = st.file_uploader("Envie o template .pptx", type=["pptx"])
    if template_file:
        forcar_redimensionamento = st.checkbox(
            f"Forçar redimensionamento para {formato_prisma}", 
            value=False,
            help="⚠️ Cuidado: isso vai esticar ou comprimir o slide. Imagens e logos podem distorcer."
        )

st.markdown("---")

# BOTÃO GERAR
if st.button("🚀 GERAR PRISMAS", type="primary", use_container_width=True):
    if not nome_col and not cargo_col and not empresa_col:
        st.error("Selecione pelo menos uma coluna para mapeamento!")
        st.stop()

    try:
        progress = st.progress(0)
        status = st.empty()

        status.info("📂 Carregando template...")

        if modo_template == "Usar template padrão":
            if not template_path or not os.path.exists(template_path):
                st.error("Template padrão não encontrado.")
                st.stop()
            prs = Presentation(template_path)
        else:
            if not template_file:
                st.error("Envie o template PowerPoint!")
                st.stop()
            template_bytes = io.BytesIO(template_file.read())
            prs = Presentation(template_bytes)
            
            # NOVIDADE: Forçar redimensionamento para template personalizado
            if forcar_redimensionamento:
                if formato_prisma == "21x14,5 cm":
                    prs.slide_width = Cm(21)
                    prs.slide_height = Cm(14.5)
                else:
                    prs.slide_width = Cm(30)
                    prs.slide_height = Cm(10.5)

        total = len(df)
        if total == 0:
            st.error("Nenhum dado no Excel!")
            st.stop()

        status.info(f"⚙️ Gerando {total} prisma(s)...")

        for i in range(total):
            duplicate_slide(prs, 0)
            progress.progress((i + 1) / (total * 2))

        campos_para_limpar = []
        if not nome_col:
            campos_para_limpar.append("NOME")
        if not cargo_col:
            campos_para_limpar.append("CARGO")
        if not empresa_col:
            campos_para_limpar.append("EMPRESA")

        for idx, (_, row) in enumerate(df.iterrows(), start=1):
            slide = prs.slides[idx]
            subs = {}
            if nome_col:
                subs["NOME"] = row[nome_col]
            if cargo_col:
                subs["CARGO"] = row[cargo_col]
            if empresa_col:
                subs["EMPRESA"] = row[empresa_col]

            substitute_text_in_slide(
                slide,
                subs,
                campos_para_limpar=campos_para_limpar,
                apagar_shape=apagar_shape_vazio,
                font_size=font_size # <-- Passando o tamanho da fonte
            )
            progress.progress((total + idx) / (total * 2))

        delete_slide(prs, 0)

        status.info("💾 Preparando arquivo PowerPoint...")
        output_pptx = io.BytesIO()
        prs.save(output_pptx)
        output_pptx.seek(0)
        pptx_bytes = output_pptx.getvalue()

        progress.progress(100)
        status.success(f"✅ {total} prisma(s) gerado(s) com sucesso!")

        st.download_button(
            label="📥 Baixar PowerPoint (.pptx)",
            data=pptx_bytes,
            file_name=f"prismas_{formato_prisma.replace(',','.').replace(' ','')}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

    except Exception as e:
        st.error(f"❌ Erro durante a geração: {e}")
        st.exception(e)