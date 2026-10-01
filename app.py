import streamlit as st
import pandas as pd
from pptx import Presentation
import io
import os
import copy
from pptx.opc.constants import RELATIONSHIP_TYPE as RT


# ============================================================
# FUNÇÕES AUXILIARES
# ============================================================

def duplicate_slide(presentation, index):
    """Duplica o slide preservando imagens, fontes, formatos e posições."""
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


def substitute_text_in_slide(slide, substituicoes, campos_para_limpar=None, apagar_shape=False):
    """
    substituicoes: dict { 'NOME': 'João', 'CARGO': 'Diretor', 'EMPRESA': 'ACME' }
    campos_para_limpar: lista de chaves cujo texto deve ser apagado do slide.
    apagar_shape: se True, remove a caixa inteira quando ela só contém a palavra-chave.
    """
    campos_para_limpar = campos_para_limpar or []
    chaves_limpar_upper = [c.upper() for c in campos_para_limpar]

    def substituir_no_run(run, subs):
        texto = run.text
        for chave, valor in subs.items():
            texto = texto.replace(chave, str(valor))
            texto = texto.replace(chave.capitalize(), str(valor))
            texto = texto.replace(chave.lower(), str(valor))
        run.text = texto

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
        # Verifica se o shape só contém palavra(s) a limpar
        if shape.has_text_frame and chaves_limpar_upper:
            texto_shape = shape.text_frame.text.upper()
            contem_limpar = any(chave in texto_shape for chave in chaves_limpar_upper)

            if contem_limpar and apagar_shape:
                texto_restante = texto_shape
                for chave in chaves_limpar_upper:
                    texto_restante = texto_restante.replace(chave, "")
                # Se só sobrou espaço/símbolos, apaga o shape inteiro
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
st.caption("Versão Web 1.2 - Criado por João Carlos")

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

# PASSO 2: Escolher aba
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

st.markdown("**Comportamento dos campos não usados**")
st.caption(
    "Se você não mapear uma coluna, os textos correspondentes (NOME, CARGO, EMPRESA) "
    "serão removidos do PPT gerado."
)
apagar_shape_vazio = st.checkbox(
    "Apagar a caixa de texto inteira (em vez de apenas o texto)",
    value=False,
    help=(
        "Marcado: se a caixa só continha 'CARGO', ela desaparece, deixando o layout mais limpo. "
        "Desmarcado: o texto é apagado mas a caixa permanece no lugar."
    ),
)

# PASSO 3: Escolha do template
st.subheader("3️⃣ Template PowerPoint")

modo_template = st.radio(
    "Como deseja escolher o template?",
    ["Usar template padrão", "Enviar meu template"],
    horizontal=True,
)

TEMPLATES_PADRAO = {
    "SESI": "templates/sesi.pptx",
    "SENAI": "templates/senai.pptx",
    "CNI": "templates/cni.pptx",
    "IEL": "templates/iel.pptx",
    "Sistema Indústria": "templates/sistema_industria.pptx",
}

template_file = None
template_path = None
nome_template = None

if modo_template == "Usar template padrão":
    nome_template = st.selectbox("Escolha o modelo", list(TEMPLATES_PADRAO.keys()))
    template_path = TEMPLATES_PADRAO[nome_template]

    if os.path.exists(template_path):
        st.success(f"✅ Template **{nome_template}** carregado")
    else:
        st.error(
            f"❌ Arquivo `{template_path}` não encontrado.\n\n"
            "Verifique se o arquivo está na pasta `templates/` do projeto."
        )
else:
    template_file = st.file_uploader("Envie o template .pptx", type=["pptx"])

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

        total = len(df)
        if total == 0:
            st.error("Nenhum dado no Excel!")
            st.stop()

        status.info(f"⚙️ Gerando {total} prisma(s)...")

        # Duplica o slide 0 (template) N vezes
        for i in range(total):
            duplicate_slide(prs, 0)
            progress.progress((i + 1) / (total * 2))

        # Descobre quais campos NÃO foram mapeados (devem ser limpos)
        campos_para_limpar = []
        if not nome_col:
            campos_para_limpar.append("NOME")
        if not cargo_col:
            campos_para_limpar.append("CARGO")
        if not empresa_col:
            campos_para_limpar.append("EMPRESA")

        # Substitui os textos nos slides duplicados
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
            )
            progress.progress((total + idx) / (total * 2))

        # Remove o slide original
        delete_slide(prs, 0)

        # Salva PPTX em memória
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
            file_name="prismas_gerados.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

    except Exception as e:
        st.error(f"❌ Erro durante a geração: {e}")
        st.exception(e)