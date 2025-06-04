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
    def reconstruct_runs_and_replace(paragraph):
        full_text = ''.join(run.text for run in paragraph.runs)
        clean_text = normalize(full_text)
        replacements = {}
        for tag, col in mapping.items():
            if col and col != "(laisser inchangée)" and col in row.index:
                placeholder = "{{" + tag + "}}"
                if normalize(placeholder) in clean_text:
                    replacements[placeholder] = str(row[col])
        if not replacements:
            return
        # Clear runs
        for run in paragraph.runs:
            run.text = ''
        # Insert new runs with original style
        new_text = full_text
        for placeholder, value in replacements.items():
            new_text = normalize(new_text).replace(normalize(placeholder), value)
        paragraph.add_run(new_text)

    def process(container):
        for p in container.paragraphs:
            reconstruct_runs_and_replace(p)
        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:
                    process(cell)

    process(template)
    for section in template.sections:
        process(section.header)
        process(section.footer)

def main():
    st.title("Publipostage Streamlit – Version 3.13.4")

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

    if word_file and excel_file:
        if df is None:
            df = pd.read_excel(excel_file)
        st.markdown("### Mappage balises → colonnes Excel")
        cols = ["(laisser inchangée)"] + list(df.columns)
        for tag in sorted(tags):
            default = cols.index(tag) if tag in df.columns else 0
            mapping[tag] = st.selectbox(f"{{{{{tag}}}}}", cols, index=default)

        if st.button("Générer les documents"):
            import io
            import zipfile
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
