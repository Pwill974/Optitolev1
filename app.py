import csv
import io
import re
import unicodedata
import zipfile
from collections import defaultdict
from numbers import Number
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

REFERENCE_HEADERS = {
    "repere", "reperepiece", "numerorepere", "numerodepiece",
    "numeropiece", "piece", "position", "mark", "partmark", "mainpartmark",
}
QUANTITY_HEADERS = {"quantite", "qte", "nombre", "nb", "quantity", "qty"}


def remove_accents(text: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFKD", text)
                   if not unicodedata.combining(c))


def cell_to_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, Number):
        number = float(value)
        if number.is_integer():
            return str(int(number))
    return str(value).strip()


def normalize_reference(value: object) -> str:
    text = remove_accents(cell_to_text(value)).upper()
    text = re.sub(r"[\s\-_]+", "", text)
    if not text:
        return ""
    if text.isdigit():
        return str(int(text))
    return text


def normalize_header(value: object) -> str:
    text = remove_accents(cell_to_text(value)).lower()
    return re.sub(r"[^a-z0-9]+", "", text)


def read_excel_rows(data: bytes):
    workbook = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    worksheet = workbook.active
    return [list(row) for row in worksheet.iter_rows(values_only=True)]


def read_csv_rows(data: bytes):
    text = None
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            pass
    if text is None:
        raise ValueError("Impossible de lire le fichier CSV.")
    try:
        dialect = csv.Sniffer().sniff(text[:4096], delimiters=";,")
    except csv.Error:
        dialect = csv.excel
        dialect.delimiter = ";"
    return [list(row) for row in csv.reader(io.StringIO(text), dialect)]


def find_columns(rows):
    for row_index, row in enumerate(rows[:30]):
        normalized = [normalize_header(value) for value in row]
        ref_i = next((i for i, value in enumerate(normalized)
                      if value in REFERENCE_HEADERS), None)
        qty_i = next((i for i, value in enumerate(normalized)
                      if value in QUANTITY_HEADERS), None)
        if ref_i is not None and qty_i is not None:
            return row_index, ref_i, qty_i
    raise ValueError("Colonnes Repère et Quantité introuvables.")


def parse_quantity(value: object, row_number: int) -> int:
    if value is None or str(value).strip() == "":
        raise ValueError(f"Quantité vide à la ligne {row_number}.")
    try:
        quantity = int(float(str(value).replace(",", ".").strip()))
    except ValueError as exc:
        raise ValueError(f"Quantité invalide à la ligne {row_number} : {value!r}") from exc
    if quantity <= 0:
        raise ValueError(f"La quantité doit être positive à la ligne {row_number}.")
    return quantity


def read_nomenclature(uploaded_file):
    data = uploaded_file.getvalue()
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix == ".xlsx":
        rows = read_excel_rows(data)
    elif suffix == ".csv":
        rows = read_csv_rows(data)
    else:
        raise ValueError("Utilisez un fichier .xlsx ou .csv.")

    header_row, ref_col, qty_col = find_columns(rows)
    result = []
    for row_number, row in enumerate(rows[header_row + 1:], start=header_row + 2):
        ref_value = row[ref_col] if ref_col < len(row) else None
        qty_value = row[qty_col] if qty_col < len(row) else None
        display_ref = cell_to_text(ref_value)
        normalized_ref = normalize_reference(ref_value)
        if not normalized_ref:
            continue
        result.append({
            "display_reference": display_ref,
            "normalized_reference": normalized_ref,
            "quantity": parse_quantity(qty_value, row_number),
        })
    if not result:
        raise ValueError("Aucune pièce exploitable trouvée.")
    return result


def read_dxf_zip(uploaded_file):
    try:
        with zipfile.ZipFile(io.BytesIO(uploaded_file.getvalue())) as archive:
            names = [name for name in archive.namelist()
                     if not name.endswith("/") and Path(name).suffix.lower() == ".dxf"]
    except zipfile.BadZipFile as exc:
        raise ValueError("Le fichier ZIP est invalide.") from exc
    if not names:
        raise ValueError("Aucun DXF trouvé dans le ZIP.")
    return sorted(names, key=str.lower)


def match_files(nomenclature, dxf_names):
    index = defaultdict(list)
    for full_name in dxf_names:
        normalized = normalize_reference(Path(full_name).stem)
        if normalized:
            index[normalized].append(full_name)

    results, used = [], set()
    for row in nomenclature:
        candidates = index.get(row["normalized_reference"], [])
        if len(candidates) == 1:
            status, dxf_name = "Trouvé", candidates[0]
            used.add(candidates[0])
        elif len(candidates) == 0:
            status, dxf_name = "DXF manquant", ""
        else:
            status, dxf_name = "Plusieurs DXF possibles", " | ".join(candidates)
            used.update(candidates)
        results.append({
            "Repère": row["display_reference"],
            "Quantité": row["quantity"],
            "Fichier DXF": dxf_name,
            "État": status,
        })
    return results, [name for name in dxf_names if name not in used]


st.set_page_config(page_title="OptiTôle Web V1", page_icon="📐", layout="wide")
st.title("📐 OptiTôle Web V1")
st.subheader("Import Advance Steel : DXF + nomenclature")
st.info("Cette V1 vérifie les correspondances entre les repères Excel et les noms des DXF.")

with st.expander("Préparer les fichiers", expanded=True):
    st.markdown("""
**DXF :** place tous les `.dxf` dans un dossier, puis compresse ce dossier en `.zip`.

**Nomenclature :** utilise un `.xlsx` ou `.csv` contenant au minimum `Repère` et `Quantité`.

Exemples acceptés : `1.dxf`, `10.dxf`, `AT 1.dxf`, `AT1.dxf`, `AT-20.dxf`.
""")

col1, col2 = st.columns(2)
with col1:
    dxf_zip = st.file_uploader("1. Charger le ZIP des DXF", type=["zip"])
with col2:
    nomenclature_file = st.file_uploader("2. Charger la nomenclature", type=["xlsx", "csv"])

if st.button("3. Analyser les correspondances", type="primary", use_container_width=True):
    if dxf_zip is None or nomenclature_file is None:
        st.warning("Charge d'abord les deux fichiers.")
    else:
        try:
            nomenclature = read_nomenclature(nomenclature_file)
            dxf_names = read_dxf_zip(dxf_zip)
            results, unused = match_files(nomenclature, dxf_names)
            df = pd.DataFrame(results)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Repères analysés", len(df))
            c2.metric("Trouvés", int((df["État"] == "Trouvé").sum()))
            c3.metric("Manquants", int((df["État"] == "DXF manquant").sum()))
            c4.metric("Ambigus", int((df["État"] == "Plusieurs DXF possibles").sum()))

            st.dataframe(df, use_container_width=True, hide_index=True)
            csv_output = df.to_csv(index=False, sep=";").encode("utf-8-sig")
            st.download_button("Télécharger le rapport CSV", csv_output,
                               "rapport_correspondances.csv", "text/csv")
            if unused:
                st.warning("DXF non utilisés :

- " + "
- ".join(unused))
            else:
                st.success("Tous les DXF ont été associés.")
        except Exception as exc:
            st.error(str(exc))

st.divider()
st.caption("V1 : contrôle des repères. Lecture géométrique et nesting ajoutés ensuite.")
