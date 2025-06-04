import streamlit as st
import re
from docx import Document
import pandas as pd
import unicodedata

def normalize(text):
    return unicodedata.normalize("NFKC", text.replace('\xa0', ' ').replace('\u200b', '')).strip()

def extract_tags_from_docx(docx_file) -> set:
    pattern = re.compile(r"\{\{\s*(.*?)\s*\}\}")
    jinja_blocks = re.compile(r"{%\s*(if|endif|for|endfor)[^%]*%}")
    doc = Document(docx_file)
    tags = set()
    jinja_found = False

    def check_text(text):
        nonlocal jinja_found
        for match in pattern.findall(text):
            tags.add(normalize(match))
        if jinja_blocks.search(text):
            jinja_found = True

    for p in doc.paragraphs:
        check_text(p.text)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                check_text(cell.text)
    for section in doc.sections:
        for part in [section.header, section.footer]:
            for p in part.paragraphs:
                check_text(p.text)
            for table in part.tables:
                for row in table.rows:
                    for cell in row.cells:
                        check_text(cell.text)

    return tags, jinja_found

def replace_placeholders_in_doc(template, mapping, row):
    def replace_in_paragraph(paragraph):
        runs = paragraph.runs
        full_text = ''.join(run.text for run in runs)
        clean_text = normalize(full_text)

        for tag, col in mapping.items():
            if not col or col == "(laisser inchangée)" or col not in row.index:
                continue
            value = str(row[col])
            regex = re.compile(r"\{\{\s*" + re.escape(tag) + r"\s*\}\}")
            match = regex.search(clean_text)
            if not match:
                continue

            # Trouver où la balise commence/termine dans les runs
            tag_start, tag_end = match.start(), match.end()
            run_positions = []
            pos = 0
            for i, run in enumerate(runs):
                text = normalize(run.text)
                if pos + len(text) >= tag_start and pos <= tag_end:
                    run_positions.append(i)
                pos += len(text)

            if run_positions:
                first_run = runs[run_positions[0]]
                for i in run_positions:
                    runs[i].text = ""
                new_run = paragraph.add_run(value)
                new_run.bold = first_run.bold
                new_run.italic = first_run.italic
                new_run.underline = first_run.underline
                new_run.font.name = first_run.font.name
                new_run.font.size = first_run.font.size

    def process(container):
        for p in container.paragraphs:
            replace_in_paragraph(p)
        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:
                    process(cell)

    process(template)
    for section in template.sections:
        process(section.header)
        process(section.footer)

def main():
    st.title("Publipostage Streamlit – Version 3.13.5")

    word_file = st.file_uploader("Modèle Word (.docx)", type="docx")
    excel_file = st.file_uploader("Fichier de données (.xls/.xlsx)", type=["xls", "xlsx"])

    mapping = {}
    tags = set()
    df = None
    jinja_found = False

    if word_file:
        tags, jinja_found = extract_tags_from_docx(word_file)

    if jinja_found:
        st.warning("⚠️ Le modèle Word contient des blocs conditionnels comme `{% if ... %}`. Ceux-ci ne seront pas traités.")

    if word_file or excel_file:
        with st.expander("📄 Afficher les détails du modèle (balises et colonnes détectées)"):
            if tags:
                st.markdown("### Balises détectées dans le modèle Word")
                for tag in sorted(tags):
                    st.write(f"- **{{{{{tag}}}}}**")
            elif word_file:
                st.info("Aucune balise {{…}} trouvée dans le document.")
            if excel_file:
                df = pd.read_excel(excel_file)
                st.markdown("### Colonnes détectées dans le fichier Excel")
                st.write(list(df.columns))

    confirmed = False
    if word_file and excel_file:
        if df is None:
            df = pd.read_excel(excel_file)
        st.markdown("### Mappage balises → colonnes Excel")
        cols = ["(laisser inchangée)"] + list(df.columns)
        for tag in sorted(tags):
            default = cols.index(tag) if tag in df.columns else 0
            mapping[tag] = st.selectbox(f"{{{{{tag}}}}}", cols, index=default)
        if st.button("Valider le mappage"):
            st.success("✅ Mappage appliqué !")
            confirmed = True

    if word_file and excel_file and mapping:
        if st.button("Générer les documents"):
            import io
            import zipfile
            df = pd.read_excel(excel_file)
            zip_io = io.BytesIO()
            with zipfile.ZipFile(zip_io, mode="w") as zf:
                for i, row in df.iterrows():
                    template = Document(word_file)
                    replace_placeholders_in_doc(template, mapping, row)
                    key = mapping.get('Name')
                    fname = str(row[key]) if key and key in row.index else str(i)
                    output_io = io.BytesIO()
                    template.save(output_io)
                    zf.writestr(f"{fname}_{i}.docx", output_io.getvalue())
            zip_io.seek(0)
            st.download_button(
                "📥 Télécharger le ZIP des documents",
                data=zip_io,
                file_name="publipostage_documents.zip",
                mime="application/zip"
            )

if __name__ == "__main__":
    main()
