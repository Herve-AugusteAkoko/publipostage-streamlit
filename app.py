import streamlit as st
import re
from docx import Document
import pandas as pd
import io
import zipfile

def extract_tags_from_docx(docx_file) -> set:
    pattern = re.compile(r"\{\{\s*(.*?)\s*\}\}")
    doc = Document(docx_file)
    tags = set()
    for p in doc.paragraphs:
        for match in pattern.findall(p.text):
            tags.add(match)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for match in pattern.findall(cell.text):
                    tags.add(match)
    for section in doc.sections:
        for part in [section.header, section.footer]:
            for p in part.paragraphs:
                for match in pattern.findall(p.text):
                    tags.add(match)
            for table in part.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for match in pattern.findall(cell.text):
                            tags.add(match)
    return tags

def replace_placeholders_in_doc(template, mapping, row):
    def process_paragraph(paragraph):
        full_text = ''.join([run.text for run in paragraph.runs])
        replacements = {}
        for tag, col in mapping.items():
            if col and col != "(laisser inchangée)" and col in row.index:
                value = str(row[col])
                placeholder = "{{" + tag + "}}"
                if placeholder in full_text:
                    replacements[placeholder] = value
        if replacements:
            for run in paragraph.runs:
                run.text = ''
            combined_text = full_text
            for placeholder, value in replacements.items():
                combined_text = combined_text.replace(placeholder, value)
            paragraph.add_run(combined_text)

    def process_container(container):
        for paragraph in container.paragraphs:
            process_paragraph(paragraph)
        for table in container.tables:
            for row in table.rows:
                for cell in row.cells:
                    process_container(cell)

    process_container(template)
    for section in template.sections:
        process_container(section.header)
        process_container(section.footer)

def main():
    st.title("Assistant de génération de documents juridiques")
    st.markdown("**Étape 1 :** Importez votre modèle Word contenant des balises comme {{Nom}}, {{Adresse}}, etc.")

    word_file = st.file_uploader("Modèle de courrier à personnaliser (.docx)", type="docx")
    excel_file = st.file_uploader("Tableur contenant les informations clients (.xls/.xlsx)", type=["xls", "xlsx"])

    mapping = {}
    tags = set()
    df = None

    if word_file or excel_file:
        with st.expander("🧾 Voir les balises détectées et colonnes disponibles"):
            if word_file:
                tags = extract_tags_from_docx(word_file)
                st.markdown("### Balises détectées dans le modèle")
                if tags:
                    for tag in sorted(tags):
                        st.write(f"- **{{{{{tag}}}}}**")
                else:
                    st.info("Aucune balise {{…}} trouvée dans le document.")
            if excel_file:
                df = pd.read_excel(excel_file)
                st.markdown("### Colonnes présentes dans le tableur")
                st.write(list(df.columns))

    if word_file and excel_file:
        if df is None:
            df = pd.read_excel(excel_file)
        if not tags:
            tags = extract_tags_from_docx(word_file)
        st.markdown("**Étape 2 :** Associez chaque balise aux colonnes du fichier Excel")
        cols = ["(laisser inchangée)"] + list(df.columns)
        for tag in sorted(tags):
            default = cols.index(tag) if tag in df.columns else 0
            mapping[tag] = st.selectbox(f"Balise : {{{{{tag}}}}}", cols, index=default)

        st.markdown("**Étape 3 :** Cliquez ci-dessous pour générer les documents personnalisés")
        if st.button("Créer les courriers personnalisés"):
            zip_io = io.BytesIO()
            with zipfile.ZipFile(zip_io, mode="w") as zf:
                for i, row in df.iterrows():
                    template = Document(word_file)
                    replace_placeholders_in_doc(template, mapping, row)
                    key = mapping.get('Nom') or mapping.get('name')
                    fname = str(row[key]) if key and key in row.index else f"document_{i}"
                    output_io = io.BytesIO()
                    template.save(output_io)
                    zf.writestr(f"{fname}.docx", output_io.getvalue())

            zip_io.seek(0)
            st.download_button(
                "📥 Télécharger tous les documents (ZIP)",
                data=zip_io,
                file_name="Documents_juridiques_personnalises.zip",
                mime="application/zip"
            )

    st.markdown("""
    ---
    🔒 *Confidentialité assurée : aucun fichier n'est stocké ou transmis. Toutes les opérations sont réalisées localement dans votre navigateur.*
    """)

if __name__ == "__main__":
    main()
