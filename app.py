import streamlit as st
import re
from docx import Document
import pandas as pd
import unicodedata
import os
import io
import zipfile

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
    """
    Remplace les placeholders {{Tag}} même s’ils sont répartis sur plusieurs runs,
    en écrivant la valeur directement dans le premier run impacté,
    et en vidant le texte des autres runs concernés.
    """
    def replace_in_paragraph(paragraph):
        runs = paragraph.runs
        if not runs:
            return

        # Construire la liste de textes de runs pour déterminer les positions
        text_runs = [run.text for run in runs]
        full_text = ''.join(text_runs)

        for tag, col in mapping.items():
            if not col or col == "(laisser inchangée)" or col not in row.index:
                continue
            value = str(row[col])
            regex = re.compile(r"\{\{\s*" + re.escape(tag) + r"\s*\}\}")

            # Tant qu’il y a une occurrence dans full_text
            while True:
                match = regex.search(full_text)
                if not match:
                    break
                start, end = match.start(), match.end()

                # Identifier les runs concernés par cette occurrence
                run_positions = []
                pos = 0
                for i, txt in enumerate(text_runs):
                    if pos + len(txt) > start and pos < end:
                        run_positions.append(i)
                    pos += len(txt)

                if not run_positions:
                    break

                # On vide tous les runs concernés
                for idx in run_positions:
                    runs[idx].text = ""

                # On écrit la valeur dans le premier run impacté
                first_idx = run_positions[0]
                runs[first_idx].text = value

                # Recalculer text_runs et full_text pour gérer plusieurs occurrences
                text_runs = [r.text for r in runs]
                full_text = ''.join(text_runs)

    def process(container):
        for p in container.paragraphs:
            replace_in_paragraph(p)
        for table in container.tables:
            for r in table.rows:
                for cell in r.cells:
                    process(cell)

    process(template)
    for section in template.sections:
        process(section.header)
        process(section.footer)

def main():
    # Page config
    st.set_page_config(
        page_title="Assistant de génération de documents juridiques",
        page_icon="⚖️"
    )
    st.title("Assistant de génération de documents juridiques")

    # Initialisation de l'état de mapping
    if "mapping_done" not in st.session_state:
        st.session_state["mapping_done"] = False

    st.markdown("""
    Génération automatisée de documents juridiques  
    à partir d’un modèle Word (.docx) et d’un fichier Excel.
    """)

    with st.expander("🔐 Politique de confidentialité"):
        st.markdown("""
        Vos fichiers ne sont jamais stockés,  
        traités uniquement pendant la session.  
        ✅ Conforme RGPD.
        """)

    # Upload des fichiers
    word_file  = st.file_uploader("📄 Modèle Word (.docx)", type="docx")
    excel_file = st.file_uploader("📊 Données Excel (.xls/.xlsx)", type=["xls", "xlsx"])

    mapping = {}
    tags = set()
    df = None
    jinja_found = False

    if word_file:
        tags, jinja_found = extract_tags_from_docx(word_file)
    if jinja_found:
        st.warning("⚠️ Blocs conditionnels Jinja non traités.")

    if word_file or excel_file:
        with st.expander("📂 Aperçu des données importées"):
            if tags:
                st.markdown("### Balises détectées")
                for tag in sorted(tags):
                    st.write(f"- **{{{{{tag}}}}}**")
            elif word_file:
                st.info("Aucune balise détectée.")
            if excel_file:
                df = pd.read_excel(excel_file)
                df.columns = df.columns.str.strip()
                st.markdown("### Colonnes Excel")
                st.write(list(df.columns))

    if word_file and excel_file:
        if df is None:
            df = pd.read_excel(excel_file)
            df.columns = df.columns.str.strip()

        st.markdown("### Mapping balises → colonnes")
        tol = st.checkbox("Mode tolérant (casse/underscores ignorés)", value=False)

        cols = ["(laisser inchangée)"] + list(df.columns)
        normalized = [c.lower().replace("_", "") for c in df.columns]

        for tag in sorted(tags):
            if tol:
                tn = tag.lower().replace("_", "")
                default = normalized.index(tn) + 1 if tn in normalized else 0
            else:
                default = cols.index(tag) if tag in df.columns else 0

            mapping[tag] = st.selectbox(f"Champ `{tag}`", cols, index=default)

        if st.button("🔗 Enregistrer les correspondances"):
            st.session_state["mapping_done"] = True
            st.success("Correspondances enregistrées.")

    if word_file and excel_file and st.session_state["mapping_done"]:
        if st.button("📂 Générer les documents"):
            df = pd.read_excel(excel_file)
            df.columns = df.columns.str.strip()

            raw        = os.path.splitext(word_file.name)[0]
            clean_name = raw.replace("_", " ")
            parts      = clean_name.split(" ", 1)
            prefix     = parts[0]; rest = parts[1] if len(parts)>1 else ""

            try:
                major, minor = prefix.split("."); minor = int(minor)
            except:
                major, minor = prefix, 0

            zip_io = io.BytesIO()
            with zipfile.ZipFile(zip_io, "w") as zf:
                # Réindexer pour avoir 0,1,2...
                df = df.reset_index(drop=True)
                for idx, row in df.iterrows():
                    template = Document(word_file)
                    replace_placeholders_in_doc(template, mapping, row)
                    seq    = minor + idx + 1
                    newpref= f"{major}.{seq}"
                    key    = next((c for t,c in mapping.items() if t.lower()=="name" and c in row.index), None)
                    person = str(row[key]).strip() if key else "inconnu"
                    fname  = f"{newpref} {rest} - {person}.docx"
                    out    = io.BytesIO()
                    template.save(out)
                    zf.writestr(fname, out.getvalue())

            zip_io.seek(0)
            st.session_state["zip_data"]     = zip_io.getvalue()
            st.session_state["zip_filename"] = f"{clean_name}.zip"

        if st.session_state.get("zip_data"):
            st.download_button(
                "📥 Télécharger vos documents personnalisés",
                data=st.session_state["zip_data"],
                file_name=st.session_state["zip_filename"],
                mime="application/zip"
            )

if __name__ == "__main__":
    main()
