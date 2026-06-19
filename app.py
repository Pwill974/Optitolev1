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
    "repere",
    "reperepiece",
    "numerorepere",
    "numerodepiece",
    "numeropiece",
    "piece",
    "position",
    "mark",
    "partmark",
    "mainpartmark",
}

QUANTITY_HEADERS = {
    "quantite",
    "qte",
    "nombre",
    "nb",
    "quantity",
    "qty",
}


def remove_accents(text: str) -> str:
    return "".join(
        char
        for char in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(char)
    )


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


def read_excel_rows(data: bytes) -> list[list[object]]:
    workbook = load_workbook(
        io.BytesIO(data),
        read_only=True,
        data_only=True,
    )
    worksheet = workbook.active
    return [list(row) for row in worksheet.iter_rows(values_only=True)]


def read_csv_rows(data: bytes) -> list[list[object]]:
    decoded_text = None

    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            decoded_text = data.decode(encoding)
            break
        except UnicodeDecodeError:
            continue

    if decoded_text is None:
        raise ValueError("Impossible de lire le fichier CSV.")

    try:
        dialect = csv.Sniffer().sniff(
            decoded_text[:4096],
            delimiters=";,",
        )
    except csv.Error:
        dialect = csv.excel
        dialect.delimiter = ";"

    return [
        list(row)
        for row in csv.reader(io.StringIO(decoded_text), dialect)
    ]


def find_columns(rows: list[list[object]]) -> tuple[int, int, int]:
    for row_index, row in enumerate(rows[:30]):
        normalized = [normalize_header(value) for value in row]

        reference_index = next(
            (
                index
                for index, value in enumerate(normalized)
                if value in REFERENCE_HEADERS
            ),
            None,
        )

        quantity_index = next(
            (
                index
                for index, value in enumerate(normalized)
                if value in QUANTITY_HEADERS
            ),
            None,
        )

        if reference_index is not None and quantity_index is not None:
            return row_index, reference_index, quantity_index

    raise ValueError(
        "Les colonnes Repère et Quantité sont introuvables. "
        "Utilisez par exemple les titres « Repère » et « Quantité »."
    )


def parse_quantity(value: object, row_number: int) -> int:
    if value is None or str(value).strip() == "":
        raise ValueError(f"Quantité vide à la ligne {row_number}.")

    try:
        quantity = int(
            float(str(value).replace(",", ".").strip())
        )
    except ValueError as exc:
        raise ValueError(
            f"Quantité invalide à la ligne {row_number} : {value!r}"
        ) from exc

    if quantity <= 0:
        raise ValueError(
            f"La quantité doit être supérieure à zéro "
            f"à la ligne {row_number}."
        )

    return quantity


def read_nomenclature(uploaded_file) -> list[dict]:
    data = uploaded_file.getvalue()
    suffix = Path(uploaded_file.name).suffix.lower()

    if suffix == ".xlsx":
        rows = read_excel_rows(data)
    elif suffix == ".csv":
        rows = read_csv_rows(data)
    else:
        raise ValueError(
            "Utilisez une nomenclature au format .xlsx ou .csv."
        )

    header_row, reference_column, quantity_column = find_columns(rows)
    result = []

    for row_number, row in enumerate(
        rows[header_row + 1 :],
        start=header_row + 2,
    ):
        reference_value = (
            row[reference_column]
            if reference_column < len(row)
            else None
        )

        quantity_value = (
            row[quantity_column]
            if quantity_column < len(row)
            else None
        )

        display_reference = cell_to_text(reference_value)
        normalized_reference = normalize_reference(reference_value)

        if not normalized_reference:
            continue

        result.append(
            {
                "display_reference": display_reference,
                "normalized_reference": normalized_reference,
                "quantity": parse_quantity(
                    quantity_value,
                    row_number,
                ),
            }
        )

    if not result:
        raise ValueError(
            "Aucune pièce exploitable n'a été trouvée."
        )

    return result


def read_dxf_zip(uploaded_file) -> list[str]:
    data = uploaded_file.getvalue()

    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            names = [
                name
                for name in archive.namelist()
                if not name.endswith("/")
                and Path(name).suffix.lower() == ".dxf"
            ]
    except zipfile.BadZipFile as exc:
        raise ValueError("Le fichier ZIP est invalide.") from exc

    if not names:
        raise ValueError(
            "Aucun fichier DXF n'a été trouvé dans le ZIP."
        )

    return sorted(names, key=str.lower)


def match_files(
    nomenclature: list[dict],
    dxf_names: list[str],
) -> tuple[list[dict], list[str]]:
    index = defaultdict(list)

    for full_name in dxf_names:
        stem = Path(full_name).stem
        normalized = normalize_reference(stem)

        if normalized:
            index[normalized].append(full_name)

    results = []
    used = set()

    for row in nomenclature:
        candidates = index.get(
            row["normalized_reference"],
            [],
        )

        if len(candidates) == 1:
            status = "Trouvé"
            dxf_name = candidates[0]
            used.add(candidates[0])

        elif len(candidates) == 0:
            status = "DXF manquant"
            dxf_name = ""

        else:
            status = "Plusieurs DXF possibles"
            dxf_name = " | ".join(candidates)
            used.update(candidates)

        results.append(
            {
                "Repère": row["display_reference"],
                "Quantité": row["quantity"],
                "Fichier DXF": dxf_name,
                "État": status,
            }
        )

    unused = [
        name
        for name in dxf_names
        if name not in used
    ]

    return results, unused


st.set_page_config(
    page_title="OptiTôle Web V1",
    page_icon="📐",
    layout="wide",
)

st.title("📐 OptiTôle Web V1")
st.subheader("Import Advance Steel : DXF + nomenclature")

st.info(
    "Cette première version vérifie les correspondances "
    "entre les repères de la nomenclature "
    "et les noms des fichiers DXF."
)

with st.expander("Préparation des fichiers", expanded=True):
    st.markdown(
        """
**1. Fichiers DXF**

Place tous les fichiers `.dxf` dans un dossier,
puis compresse ce dossier au format `.zip`.

Exemples de noms acceptés :

- `1.dxf`
- `10.dxf`
- `AT 1.dxf`
- `AT1.dxf`
- `AT-20.dxf`

**2. Nomenclature**

Utilise un fichier `.xlsx` ou `.csv`
avec au minimum les colonnes :

| Repère | Quantité |
|---|---:|
| 1 | 5 |
| 10 | 2 |
| AT 1 | 4 |
        """
    )

column_1, column_2 = st.columns(2)

with column_1:
    dxf_zip = st.file_uploader(
        "1. Charger le ZIP contenant les DXF",
        type=["zip"],
    )

with column_2:
    nomenclature_file = st.file_uploader(
        "2. Charger la nomenclature Excel ou CSV",
        type=["xlsx", "csv"],
    )

analyze = st.button(
    "3. Analyser les correspondances",
    type="primary",
    use_container_width=True,
)

if analyze:
    if dxf_zip is None or nomenclature_file is None:
        st.warning(
            "Charge d'abord le ZIP des DXF "
            "et la nomenclature."
        )

    else:
        try:
            nomenclature = read_nomenclature(
                nomenclature_file
            )

            dxf_names = read_dxf_zip(dxf_zip)

            results, unused = match_files(
                nomenclature,
                dxf_names,
            )

            dataframe = pd.DataFrame(results)

            found = int(
                (dataframe["État"] == "Trouvé").sum()
            )

            missing = int(
                (dataframe["État"] == "DXF manquant").sum()
            )

            ambiguous = int(
                (
                    dataframe["État"]
                    == "Plusieurs DXF possibles"
                ).sum()
            )

            metric_1, metric_2, metric_3, metric_4 = (
                st.columns(4)
            )

            metric_1.metric(
                "Repères analysés",
                len(dataframe),
            )

            metric_2.metric(
                "Trouvés",
                found,
            )

            metric_3.metric(
                "Manquants",
                missing,
            )

            metric_4.metric(
                "Ambigus",
                ambiguous,
            )

            st.dataframe(
                dataframe,
                use_container_width=True,
                hide_index=True,
            )

            csv_output = dataframe.to_csv(
                index=False,
                sep=";",
            ).encode("utf-8-sig")

            st.download_button(
                "Télécharger le rapport CSV",
                data=csv_output,
                file_name="rapport_correspondances.csv",
                mime="text/csv",
            )

            if unused:
                unused_text = "\n- ".join(unused)
                st.warning(
                    "DXF non utilisés :\n\n- "
                    + unused_text
                )
            else:
                st.success(
                    "Tous les DXF ont été associés "
                    "à la nomenclature."
                )

        except Exception as exc:
            st.error(str(exc))

st.divider()

st.caption(
    "V1 : contrôle des repères uniquement. "
    "La lecture géométrique des DXF "
    "et le nesting seront ajoutés ensuite."
)
