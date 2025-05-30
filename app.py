import streamlit as st
import re
from docx import Document
import pandas as pd

def extract_tags_from_docx(docx_file) -> set:
    """
    Lit un .docx et renvoie un set de toutes les balises {{Tag}} sans espaces.
    """
    pattern = re.compile(r"\{\{\s*(.*?)\s*\}\}")
    doc = Document(docx_file)
    tags = set()
    # Paragraphes et tables du corps
    for p in doc.paragraphs:
        for match in pattern.findall(p.text):
            tags.add(match)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for match in pattern.findall(cell.text):
                    tags.add(match)
    # Paragraphes et tables des en-têtes et pieds-de-page
    for section in doc.sections:
        header = section.header
        for p in header.paragraphs:
            for match in pattern.findall(p.text):
                tags.add(match)
        for table in header.tables:
            for row in table.rows:
                for cell in row.cells:
                    for match in pattern.findall(cell.text):
                        tags.add(match)
        footer = section.footer
        for p in footer.paragraphs:
            for match in pattern.findall(p.text):
                tags.add(match)
        for table in footer.tables:
            for row in table.rows:
                for cell in row.cells:
                    for match in pattern.findall(cell.text):
                        tags.add(match)
    return tags

def replace_placeholders_in_doc(template, mapping, row):
    """
    Remplace les placeholders dans le document (body, header, footer) selon mapping.
    mapping: dict tag->colname
    row: Series pandas
    """
    def replace_in_paragraphs(paragraphs):
        for p in paragraphs:
            full_text = "".join([run.text for run in p.runs])
            for tag, col in mapping.items():
                if col and col != "(laisser inchangée)" and col in row.index:
                    placeholder_regex = re.compile(r"\{\{\s*" + re.escape(tag) + r"\s*\}\}")
                    new_text = placeholder_regex.sub(str(row[col]), full_text)
                    if new_text != full_text:
                        for run in p.runs:
                            p._p.remove(run._r)
                        p.add_run(new_text)
                        full_text = new_text
    # Remplacement dans le corps
    replace_in_paragraphs(template.paragraphs)
    for table in template.tables:
        for row_cells in table.rows:
            for cell in row_cells.cells:
                replace_in_paragraphs(cell.paragraphs)
    # Remplacement dans en-têtes et pieds-de-page
    for section in template.sections:
        replace_in_paragraphs(section.header.paragraphs)
        for table in section.header.tables:
            for row_cells in table.rows:
                for cell in row_cells.cells:
                    replace_in_paragraphs(cell.paragraphs)
        replace_in_paragraphs(section.footer.paragraphs)
        for table in section.footer.tables:
            for row_cells in table.rows:
                for cell in row_cells.cells:
                    replace_in_paragraphs(cell.paragraphs)

def main():
    st.title("Publipostage Streamlit")
    st.markdown("**Étape 1 :** Sélectionnez votre modèle Word et votre fichier Excel.")

    word_file = st.file_uploader("Modèle Word (.docx)", type="docx")
    excel_file = st.file_uploader("Fichier de données (.xls/.xlsx)", type=["xls", "xlsx"])

    mapping = {}
    tags = set()
    df = None

    # Affichage conditionnel dans un expander
    if word_file or excel_file:
        with st.expander("📄 Afficher les détails du modèle (balises et colonnes détectées)"):
            if word_file:
                tags = extract_tags_from_docx(word_file)
                st.markdown("### Balises détectées dans le modèle Word")
                if tags:
                    for tag in sorted(tags):
                        st.write(f"- **{{{{{tag}}}}}**")
                else:
                    st.info("Aucune balise {{…}} trouvée dans le document.")

            if excel_file:
                df = pd.read_excel(excel_file)
                st.markdown("### Colonnes détectées dans le fichier Excel")
                st.write(list(df.columns))

    # Mappage balise → colonne
    if word_file and excel_file:
        if df is None:
            df = pd.read_excel(excel_file)
        if not tags:
            tags = extract_tags_from_docx(word_file)
        st.markdown("### Mappage balises → colonnes Excel")
        cols = ["(laisser inchangée)"] + list(df.columns)
        for tag in sorted(tags):
            default = cols.index(tag) if tag in df.columns else 0
            mapping[tag] = st.selectbox(f"{{{{{tag}}}}}", cols, index=default)
        if st.button("Valider le mappage"):
            st.success("Mappage enregistré !")

    # Génération et téléchargement
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
                "Télécharger le ZIP des documents",
                data=zip_io,
                file_name="publipostage_documents.zip",
                mime="application/zip"
            )

if __name__ == "__main__":
    main()
