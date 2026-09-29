import streamlit as st
import pandas as pd
from pptx import Presentation
from copy import deepcopy
import io
import copy
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

def duplicate_slide(presentation, index):
    """
    Duplica o slide no índice informado preservando imagens, fontes,
    formatos, posições, tabelas, gráficos, etc. — de forma limpa, sem
    deixar XML inconsistente que force o PowerPoint a reparar.
    """
    source = presentation.slides[index]

    # 1. Cria o novo slide com o MESMO layout do original
    dest = presentation.slides.add_slide(source.slide_layout)

    # 2. Remove todos os shapes herdados do layout
    for shp in list(dest.shapes):
        shp._element.getparent().remove(shp._element)

    # 3. Copia o XML do <p:spTree> do original, mas com filtros
    src_spTree = source.shapes._spTree
    dst_spTree = dest.shapes._spTree

    # Remove os elementos "estruturais" do spTree de destino,
    # mantendo apenas nvGrpSpPr e grpSpPr
    for child in list(dst_spTree):
        tag = child.tag.split('}')[-1]
        if tag not in ('nvGrpSpPr', 'grpSpPr'):
            dst_spTree.remove(child)

    # Copia os shapes (sp, pic, graphicFrame, grpSp, etc.) do original
    for child in list(src_spTree):
        tag = child.tag.split('}')[-1]
        if tag in ('nvGrpSpPr', 'grpSpPr'):
            continue  # pula os elementos estruturais
        new_el = copy.deepcopy(child)
        dst_spTree.append(new_el)

    # 4. Remapeia os rId (imagens, hyperlinks, vídeos, gráficos, etc.)
    rels_map = {}
    for rel in source.part.rels.values():
        # Pula o slideLayout — o add_slide já criou o relacionamento
        if rel.reltype == RT.SLIDE_LAYOUT:
            continue
        # Pula notesSlide — não queremos herdar notas
        if 'notesSlide' in rel.reltype:
            continue

        if rel.is_external:
            new_rid = dest.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
        else:
            new_rid = dest.part.relate_to(rel.target_part, rel.reltype)
        rels_map[rel.rId] = new_rid

    # 5. Reescreve TODOS os atributos r:embed / r:link / r:id no XML do novo slide
    for elem in dst_spTree.iter():
        for attr in list(elem.attrib.keys()):
            if attr.endswith('}embed') or attr.endswith('}link') or attr.endswith('}id'):
                old_rid = elem.get(attr)
                if old_rid and old_rid in rels_map:
                    elem.set(attr, rels_map[old_rid])

    return dest


def delete_slide(presentation, index):
    """Remove o slide no índice especificado."""
    xml_slides = presentation.slides._sldIdLst
    slides = list(xml_slides)
    xml_slides.remove(slides[index])


def substitute_text_in_slide(slide, substituicoes):
    """
    substituicoes: dict { 'NOME': 'João', 'CARGO': 'Diretor', 'EMPRESA': 'ACME' }
    Percorre shapes (incluindo tabelas e grupos) e substitui.
    """
    def process_text_frame(tf):
        for paragraph in tf.paragraphs:
            for run in paragraph.runs:
                texto = run.text
                for chave, valor in substituicoes.items():
                    texto = texto.replace(chave, str(valor))
                    texto = texto.replace(chave.capitalize(), str(valor))
                    texto = texto.replace(chave.lower(), str(valor))
                run.text = texto

    def process_shape(shape):
        if shape.has_text_frame:
            process_text_frame(shape.text_frame)
        if shape.has_table:
            for row in shape.table.rows:
                for cell in row.cells:
                    process_text_frame(cell.text_frame)
        if shape.shape_type == 6:  # GROUP
            for sub in shape.shapes:
                process_shape(sub)

    for shape in slide.shapes:
        process_shape(shape)


# ============================================================
# INTERFACE STREAMLIT
# ============================================================

st.set_page_config(page_title="Gerador de Prismas", page_icon="🔷", layout="centered")

st.title("🔷 Gerador de Prismas")
st.caption("Versão Web 1.0 - Criado por João Carlos")

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

# PASSO 3: Upload do template
st.subheader("3️⃣ Template PowerPoint")
template_file = st.file_uploader("Envie o template .pptx", type=["pptx"])

st.markdown("---")

# BOTÃO GERAR
if st.button("🚀 GERAR PRISMAS", type="primary", use_container_width=True):
    if not nome_col and not cargo_col and not empresa_col:
        st.error("Selecione pelo menos uma coluna para mapeamento!")
        st.stop()
    if not template_file:
        st.error("Envie o template PowerPoint!")
        st.stop()

    try:
        progress = st.progress(0)
        status = st.empty()

        status.info("📂 Carregando template...")
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
            substitute_text_in_slide(slide, subs)
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

        # Botão de download PPTX
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