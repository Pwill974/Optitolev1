import csv
import io
import math
import re
import tempfile
import unicodedata
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from numbers import Number
from pathlib import Path
from typing import Iterable

import ezdxf
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from ezdxf.path import make_path
from openpyxl import load_workbook
from shapely import affinity
from shapely.geometry import LinearRing, Polygon, box
from shapely.ops import unary_union
from shapely.validation import make_valid


# ============================================================
# Modèles de données
# ============================================================

@dataclass
class NomenclatureItem:
    reference_display: str
    reference_key: str
    quantity: int
    thickness: str
    material: str


@dataclass
class DxfPiece:
    reference_display: str
    reference_key: str
    source_name: str
    polygon: Polygon
    quantity: int
    thickness: str
    material: str


@dataclass
class Placement:
    reference: str
    source_name: str
    sheet_index: int
    rotation: int
    polygon: Polygon
    original_polygon: Polygon
    copy_index: int
    thickness: str
    material: str


# ============================================================
# Normalisation
# ============================================================

REFERENCE_HEADERS = {
    "repere", "reperepiece", "numerorepere", "numerodepiece",
    "numeropiece", "piece", "position", "mark", "partmark",
    "mainpartmark",
}
QUANTITY_HEADERS = {
    "quantite", "qte", "nombre", "nb", "quantity", "qty",
}
THICKNESS_HEADERS = {
    "epaisseur", "ep", "thickness", "plate thickness", "thk",
}
MATERIAL_HEADERS = {
    "matiere", "materiau", "material", "nuance", "grade", "steelgrade",
}


def remove_accents(text: str) -> str:
    return "".join(
        char for char in unicodedata.normalize("NFKD", text)
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


def reference_from_dxf_filename(filename: str) -> str:
    """
    Extrait le repère réel depuis un nom de fichier Advance Steel.

    Exemples :
        NC/13.nc.dxf          -> 13
        NC/AT1.nc.dxf         -> AT1
        NC/AT1.nc.err.dxf     -> AT1
        NC/AT1A11.nc.err.dxf  -> AT1A11
        10.dxf                -> 10
    """
    filename_only = Path(filename).name

    # Retirer l'extension finale .dxf.
    cleaned = re.sub(r"(?i)\.dxf$", "", filename_only).strip()

    # Retirer successivement les suffixes techniques ajoutés par Advance Steel.
    technical_suffixes = ("nc", "err", "dstv", "cnc", "cam")

    while True:
        previous = cleaned
        suffix_pattern = r"(?i)\.(?:" + "|".join(technical_suffixes) + r")$"
        cleaned = re.sub(suffix_pattern, "", cleaned).strip()

        if cleaned == previous:
            break

    return normalize_reference(cleaned)


def normalize_group_value(value: object, fallback: str) -> str:
    text = cell_to_text(value)
    return text if text else fallback


# ============================================================
# Lecture de la nomenclature
# ============================================================

def read_excel_rows(data: bytes) -> list[list[object]]:
    """
    Recherche automatiquement une feuille Excel contenant les colonnes
    Repère et Quantité. Cela évite l'erreur lorsque la feuille active du
    classeur est absente, masquée ou invalide.
    """
    try:
        workbook = load_workbook(
            io.BytesIO(data),
            read_only=True,
            data_only=True,
        )
    except Exception as exc:
        raise ValueError(
            "Impossible d'ouvrir le fichier Excel. "
            "Vérifiez qu'il s'agit bien d'un fichier .xlsx valide."
        ) from exc

    worksheets = list(workbook.worksheets)

    if not worksheets:
        workbook.close()
        raise ValueError(
            "Le fichier Excel ne contient aucune feuille de calcul exploitable."
        )

    first_non_empty_rows = None

    try:
        for worksheet in worksheets:
            rows = [
                list(row)
                for row in worksheet.iter_rows(values_only=True)
            ]

            if not rows or not any(
                any(cell not in (None, "") for cell in row)
                for row in rows
            ):
                continue

            if first_non_empty_rows is None:
                first_non_empty_rows = rows

            try:
                find_header_row(rows)
                return rows
            except ValueError:
                continue
    finally:
        workbook.close()

    if first_non_empty_rows is not None:
        return first_non_empty_rows

    raise ValueError(
        "Toutes les feuilles du fichier Excel sont vides."
    )


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
        dialect = csv.Sniffer().sniff(decoded_text[:4096], delimiters=";,")
    except csv.Error:
        dialect = csv.excel
        dialect.delimiter = ";"

    return [list(row) for row in csv.reader(io.StringIO(decoded_text), dialect)]


def find_column(normalized_headers: list[str], accepted: set[str]) -> int | None:
    return next(
        (index for index, value in enumerate(normalized_headers) if value in accepted),
        None,
    )


def find_header_row(rows: list[list[object]]) -> tuple[int, dict[str, int | None]]:
    for row_index, row in enumerate(rows[:40]):
        normalized = [normalize_header(value) for value in row]
        reference_index = find_column(normalized, REFERENCE_HEADERS)
        quantity_index = find_column(normalized, QUANTITY_HEADERS)

        if reference_index is not None and quantity_index is not None:
            return row_index, {
                "reference": reference_index,
                "quantity": quantity_index,
                "thickness": find_column(normalized, THICKNESS_HEADERS),
                "material": find_column(normalized, MATERIAL_HEADERS),
            }

    raise ValueError(
        "Les colonnes Repère et Quantité sont introuvables. "
        "Utilisez au minimum les titres « Repère » et « Quantité »."
    )


def parse_quantity(value: object, row_number: int) -> int:
    if value is None or str(value).strip() == "":
        raise ValueError(f"Quantité vide à la ligne {row_number}.")
    try:
        quantity = int(float(str(value).replace(",", ".").strip()))
    except ValueError as exc:
        raise ValueError(
            f"Quantité invalide à la ligne {row_number} : {value!r}"
        ) from exc
    if quantity <= 0:
        raise ValueError(
            f"La quantité doit être supérieure à zéro à la ligne {row_number}."
        )
    return quantity


def read_nomenclature(uploaded_file) -> list[NomenclatureItem]:
    data = uploaded_file.getvalue()
    suffix = Path(uploaded_file.name).suffix.lower()

    if suffix == ".xlsx":
        rows = read_excel_rows(data)
    elif suffix == ".csv":
        rows = read_csv_rows(data)
    else:
        raise ValueError("Utilisez une nomenclature au format .xlsx ou .csv.")

    header_row, columns = find_header_row(rows)
    result: list[NomenclatureItem] = []

    for row_number, row in enumerate(rows[header_row + 1:], start=header_row + 2):
        def get_cell(index: int | None):
            if index is None or index >= len(row):
                return None
            return row[index]

        reference_value = get_cell(columns["reference"])
        reference_display = cell_to_text(reference_value)
        reference_key = normalize_reference(reference_value)

        if not reference_key:
            continue

        result.append(
            NomenclatureItem(
                reference_display=reference_display,
                reference_key=reference_key,
                quantity=parse_quantity(get_cell(columns["quantity"]), row_number),
                thickness=normalize_group_value(
                    get_cell(columns["thickness"]), "Non renseignée"
                ),
                material=normalize_group_value(
                    get_cell(columns["material"]), "Non renseignée"
                ),
            )
        )

    if not result:
        raise ValueError("Aucune pièce exploitable n'a été trouvée.")

    return result


# ============================================================
# Lecture des DXF
# ============================================================

def flatten_entity(entity, tolerance: float) -> list[tuple[float, float]]:
    """
    Convertit une entité fermée DXF en suite de points.
    Les arcs et courbes sont approchés par segments.
    """
    try:
        path = make_path(entity)
        points = [(float(v.x), float(v.y)) for v in path.flattening(tolerance)]
        return points
    except Exception:
        return []


def clean_ring(points: list[tuple[float, float]]) -> list[tuple[float, float]]:
    cleaned: list[tuple[float, float]] = []
    for point in points:
        if not cleaned or point != cleaned[-1]:
            cleaned.append(point)

    if len(cleaned) >= 2 and cleaned[0] == cleaned[-1]:
        cleaned.pop()

    return cleaned


def entity_is_closed(entity) -> bool:
    entity_type = entity.dxftype()

    if entity_type == "LWPOLYLINE":
        return bool(entity.closed)

    if entity_type == "POLYLINE":
        return bool(entity.is_closed)

    if entity_type in {"CIRCLE", "ELLIPSE"}:
        return True

    if entity_type == "SPLINE":
        return bool(getattr(entity, "closed", False))

    return False


def polygons_from_dxf_bytes(data: bytes, source_name: str, tolerance: float) -> list[Polygon]:
    with tempfile.NamedTemporaryFile(suffix=".dxf", delete=False) as tmp:
        tmp.write(data)
        temp_name = tmp.name

    try:
        document = ezdxf.readfile(temp_name)
    except Exception as exc:
        raise ValueError(f"DXF illisible : {source_name}") from exc
    finally:
        try:
            Path(temp_name).unlink(missing_ok=True)
        except Exception:
            pass

    modelspace = document.modelspace()
    candidate_polygons: list[Polygon] = []

    for entity in modelspace:
        if not entity_is_closed(entity):
            continue

        points = clean_ring(flatten_entity(entity, tolerance))
        if len(points) < 3:
            continue

        try:
            ring = LinearRing(points)
            polygon = Polygon(ring)
        except Exception:
            continue

        if polygon.is_empty:
            continue

        if not polygon.is_valid:
            polygon = make_valid(polygon)

        if polygon.geom_type == "MultiPolygon":
            candidate_polygons.extend(
                part for part in polygon.geoms if part.area > 0.01
            )
        elif polygon.geom_type == "Polygon" and polygon.area > 0.01:
            candidate_polygons.append(polygon)

    if not candidate_polygons:
        raise ValueError(
            f"Aucun contour fermé exploitable dans {source_name}. "
            "Vérifiez que les contours sont des polylignes fermées."
        )

    candidate_polygons.sort(key=lambda poly: poly.area, reverse=True)
    outer = candidate_polygons[0]
    holes = []

    for candidate in candidate_polygons[1:]:
        representative = candidate.representative_point()
        if outer.contains(representative):
            holes.append(candidate.exterior.coords[:])

    final_polygon = Polygon(outer.exterior.coords[:], holes)

    if not final_polygon.is_valid:
        final_polygon = make_valid(final_polygon)

    if final_polygon.geom_type == "MultiPolygon":
        final_polygon = max(final_polygon.geoms, key=lambda poly: poly.area)

    min_x, min_y, _, _ = final_polygon.bounds
    final_polygon = affinity.translate(final_polygon, xoff=-min_x, yoff=-min_y)

    return [final_polygon]


def read_dxf_zip(uploaded_file) -> dict[str, tuple[str, bytes]]:
    data = uploaded_file.getvalue()

    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            result: dict[str, tuple[str, bytes]] = {}
            for name in archive.namelist():
                if name.endswith("/") or Path(name).suffix.lower() != ".dxf":
                    continue

                key = reference_from_dxf_filename(name)
                if not key:
                    continue

                if key in result:
                    previous_name = result[key][0]
                    raise ValueError(
                        f"Deux DXF correspondent au même repère : "
                        f"{previous_name} et {name}."
                    )

                result[key] = (name, archive.read(name))
    except zipfile.BadZipFile as exc:
        raise ValueError("Le fichier ZIP des DXF est invalide.") from exc

    if not result:
        raise ValueError("Aucun fichier DXF n'a été trouvé dans le ZIP.")

    return result


def build_pieces(
    nomenclature: list[NomenclatureItem],
    dxf_files: dict[str, tuple[str, bytes]],
    tolerance: float,
) -> tuple[list[DxfPiece], list[str], list[str]]:
    pieces: list[DxfPiece] = []
    missing: list[str] = []
    used_keys: set[str] = set()

    for item in nomenclature:
        dxf_info = dxf_files.get(item.reference_key)
        if dxf_info is None:
            missing.append(item.reference_display)
            continue

        source_name, data = dxf_info
        polygons = polygons_from_dxf_bytes(data, source_name, tolerance)
        polygon = polygons[0]

        pieces.append(
            DxfPiece(
                reference_display=item.reference_display,
                reference_key=item.reference_key,
                source_name=source_name,
                polygon=polygon,
                quantity=item.quantity,
                thickness=item.thickness,
                material=item.material,
            )
        )
        used_keys.add(item.reference_key)

    unused = [
        source_name for key, (source_name, _) in dxf_files.items()
        if key not in used_keys
    ]

    return pieces, missing, unused


# ============================================================
# Moteur de nesting heuristique
# ============================================================

def rotated_at_origin(polygon: Polygon, angle: int) -> Polygon:
    rotated = affinity.rotate(polygon, angle, origin=(0, 0), use_radians=False)
    min_x, min_y, _, _ = rotated.bounds
    return affinity.translate(rotated, xoff=-min_x, yoff=-min_y)


def candidate_positions(
    placed_polygons: list[Polygon],
    margin: float,
) -> list[tuple[float, float]]:
    xs = {margin}
    ys = {margin}

    for polygon in placed_polygons:
        _, _, max_x, max_y = polygon.bounds
        xs.add(round(max_x, 6))
        ys.add(round(max_y, 6))

    candidates = [(x, y) for y in sorted(ys) for x in sorted(xs)]
    return candidates


def fits_on_sheet(
    polygon: Polygon,
    placed_polygons: list[Polygon],
    sheet_inner: Polygon,
    clearance: float,
) -> bool:
    if not sheet_inner.covers(polygon):
        return False

    for other in placed_polygons:
        if polygon.buffer(clearance / 2.0).intersects(
            other.buffer(clearance / 2.0)
        ):
            return False

    return True


def placement_score(polygon: Polygon, placed_polygons: list[Polygon]) -> tuple:
    min_x, min_y, max_x, max_y = polygon.bounds

    if placed_polygons:
        merged_bounds = unary_union(placed_polygons + [polygon]).bounds
        _, _, merged_max_x, merged_max_y = merged_bounds
    else:
        merged_max_x, merged_max_y = max_x, max_y

    return (
        round(merged_max_y, 6),
        round(merged_max_x, 6),
        round(min_y, 6),
        round(min_x, 6),
    )


def expand_piece_copies(pieces: list[DxfPiece]) -> list[tuple[DxfPiece, int]]:
    expanded: list[tuple[DxfPiece, int]] = []

    for piece in pieces:
        for copy_index in range(1, piece.quantity + 1):
            expanded.append((piece, copy_index))

    expanded.sort(
        key=lambda item: (
            item[0].polygon.area,
            max(
                item[0].polygon.bounds[2] - item[0].polygon.bounds[0],
                item[0].polygon.bounds[3] - item[0].polygon.bounds[1],
            ),
        ),
        reverse=True,
    )

    return expanded


def nest_pieces(
    pieces: list[DxfPiece],
    sheet_width: float,
    sheet_height: float,
    margin: float,
    clearance: float,
    rotations: list[int],
) -> tuple[list[Placement], int, list[str]]:
    if sheet_width <= 0 or sheet_height <= 0:
        raise ValueError("Les dimensions de la tôle doivent être positives.")

    if margin < 0 or clearance < 0:
        raise ValueError("La marge et l'espacement ne peuvent pas être négatifs.")

    inner_width = sheet_width - 2 * margin
    inner_height = sheet_height - 2 * margin

    if inner_width <= 0 or inner_height <= 0:
        raise ValueError("La marge est trop grande pour le format de tôle.")

    sheet_inner = box(
        margin,
        margin,
        sheet_width - margin,
        sheet_height - margin,
    )

    sheets: list[list[Placement]] = []
    unplaced: list[str] = []

    for piece, copy_index in expand_piece_copies(pieces):
        placed_successfully = False

        for sheet_index, sheet_placements in enumerate(sheets):
            current_polygons = [placement.polygon for placement in sheet_placements]
            best = None

            for angle in rotations:
                rotated = rotated_at_origin(piece.polygon, angle)
                width = rotated.bounds[2] - rotated.bounds[0]
                height = rotated.bounds[3] - rotated.bounds[1]

                if width > inner_width + 1e-6 or height > inner_height + 1e-6:
                    continue

                for x, y in candidate_positions(current_polygons, margin):
                    candidate = affinity.translate(rotated, xoff=x, yoff=y)

                    if not fits_on_sheet(
                        candidate,
                        current_polygons,
                        sheet_inner,
                        clearance,
                    ):
                        continue

                    score = placement_score(candidate, current_polygons)
                    if best is None or score < best[0]:
                        best = (score, angle, candidate)

            if best is not None:
                _, angle, candidate = best
                sheet_placements.append(
                    Placement(
                        reference=piece.reference_display,
                        source_name=piece.source_name,
                        sheet_index=sheet_index,
                        rotation=angle,
                        polygon=candidate,
                        original_polygon=piece.polygon,
                        copy_index=copy_index,
                        thickness=piece.thickness,
                        material=piece.material,
                    )
                )
                placed_successfully = True
                break

        if placed_successfully:
            continue

        # Créer une nouvelle tôle
        best = None
        empty_polygons: list[Polygon] = []

        for angle in rotations:
            rotated = rotated_at_origin(piece.polygon, angle)
            width = rotated.bounds[2] - rotated.bounds[0]
            height = rotated.bounds[3] - rotated.bounds[1]

            if width > inner_width + 1e-6 or height > inner_height + 1e-6:
                continue

            candidate = affinity.translate(rotated, xoff=margin, yoff=margin)

            if fits_on_sheet(candidate, empty_polygons, sheet_inner, clearance):
                score = placement_score(candidate, empty_polygons)
                if best is None or score < best[0]:
                    best = (score, angle, candidate)

        if best is None:
            unplaced.append(
                f"{piece.reference_display} - copie {copy_index}"
            )
            continue

        _, angle, candidate = best
        new_sheet_index = len(sheets)
        sheets.append(
            [
                Placement(
                    reference=piece.reference_display,
                    source_name=piece.source_name,
                    sheet_index=new_sheet_index,
                    rotation=angle,
                    polygon=candidate,
                    original_polygon=piece.polygon,
                    copy_index=copy_index,
                    thickness=piece.thickness,
                    material=piece.material,
                )
            ]
        )

    placements = [
        placement
        for sheet in sheets
        for placement in sheet
    ]

    return placements, len(sheets), unplaced


# ============================================================
# Résultats et exports
# ============================================================

def sheet_statistics(
    placements: list[Placement],
    sheet_count: int,
    sheet_width: float,
    sheet_height: float,
) -> pd.DataFrame:
    rows = []

    for sheet_index in range(sheet_count):
        sheet_placements = [
            item for item in placements if item.sheet_index == sheet_index
        ]
        used_area = sum(item.original_polygon.area for item in sheet_placements)
        sheet_area = sheet_width * sheet_height
        usage = used_area / sheet_area * 100 if sheet_area else 0

        rows.append(
            {
                "Tôle": sheet_index + 1,
                "Nombre de pièces": len(sheet_placements),
                "Surface pièces (mm²)": round(used_area, 1),
                "Utilisation (%)": round(usage, 2),
                "Chute (%)": round(100 - usage, 2),
            }
        )

    return pd.DataFrame(rows)


def placement_table(placements: list[Placement]) -> pd.DataFrame:
    rows = []

    for item in placements:
        min_x, min_y, max_x, max_y = item.polygon.bounds
        rows.append(
            {
                "Tôle": item.sheet_index + 1,
                "Repère": item.reference,
                "Copie": item.copy_index,
                "Rotation (°)": item.rotation,
                "X min (mm)": round(min_x, 2),
                "Y min (mm)": round(min_y, 2),
                "Largeur occupée (mm)": round(max_x - min_x, 2),
                "Hauteur occupée (mm)": round(max_y - min_y, 2),
                "Épaisseur": item.thickness,
                "Matière": item.material,
            }
        )

    return pd.DataFrame(rows)


def plot_sheet(
    placements: list[Placement],
    sheet_index: int,
    sheet_width: float,
    sheet_height: float,
):
    fig, ax = plt.subplots(figsize=(12, 6))
    ax.set_xlim(0, sheet_width)
    ax.set_ylim(0, sheet_height)
    ax.set_aspect("equal", adjustable="box")
    ax.set_title(f"Tôle {sheet_index + 1} — {sheet_width:g} × {sheet_height:g} mm")
    ax.set_xlabel("X (mm)")
    ax.set_ylabel("Y (mm)")
    ax.grid(True, linewidth=0.3)

    sheet_items = [
        item for item in placements if item.sheet_index == sheet_index
    ]

    for item in sheet_items:
        x, y = item.polygon.exterior.xy
        ax.fill(x, y, alpha=0.25)
        ax.plot(x, y, linewidth=1)

        for interior in item.polygon.interiors:
            hx, hy = interior.xy
            ax.plot(hx, hy, linewidth=1)

        point = item.polygon.representative_point()
        ax.text(
            point.x,
            point.y,
            f"{item.reference}\n#{item.copy_index}",
            ha="center",
            va="center",
            fontsize=7,
        )

    return fig


def add_polygon_to_dxf(modelspace, polygon: Polygon, reference: str):
    exterior = [(float(x), float(y)) for x, y in polygon.exterior.coords]
    modelspace.add_lwpolyline(exterior, close=True)

    for interior in polygon.interiors:
        hole = [(float(x), float(y)) for x, y in interior.coords]
        modelspace.add_lwpolyline(hole, close=True)

    point = polygon.representative_point()
    modelspace.add_text(
        reference,
        dxfattribs={"height": 8.0},
    ).set_placement((point.x, point.y))


def export_sheets_zip(
    placements: list[Placement],
    sheet_count: int,
    sheet_width: float,
    sheet_height: float,
) -> bytes:
    output = io.BytesIO()

    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for sheet_index in range(sheet_count):
            document = ezdxf.new("R2010")
            modelspace = document.modelspace()

            modelspace.add_lwpolyline(
                [
                    (0, 0),
                    (sheet_width, 0),
                    (sheet_width, sheet_height),
                    (0, sheet_height),
                ],
                close=True,
            )

            for item in placements:
                if item.sheet_index == sheet_index:
                    add_polygon_to_dxf(
                        modelspace,
                        item.polygon,
                        f"{item.reference}-{item.copy_index}",
                    )

            with tempfile.NamedTemporaryFile(
                suffix=".dxf",
                delete=False,
            ) as tmp:
                temp_name = tmp.name

            document.saveas(temp_name)
            archive.write(
                temp_name,
                arcname=f"tole_{sheet_index + 1}.dxf",
            )
            Path(temp_name).unlink(missing_ok=True)

    output.seek(0)
    return output.getvalue()


# ============================================================
# Interface Streamlit
# ============================================================

st.set_page_config(
    page_title="OptiTôle Pro",
    page_icon="📐",
    layout="wide",
)

st.title("📐 OptiTôle Pro v3")
st.caption(
    "Nesting de tôles à partir de DXF Advance Steel et d'une nomenclature Excel/CSV."
)

with st.sidebar:
    st.header("Paramètres de la tôle")

    sheet_width = st.number_input(
        "Largeur de la tôle (mm)",
        min_value=100.0,
        value=3000.0,
        step=100.0,
    )

    sheet_height = st.number_input(
        "Hauteur de la tôle (mm)",
        min_value=100.0,
        value=1500.0,
        step=100.0,
    )

    margin = st.number_input(
        "Marge extérieure (mm)",
        min_value=0.0,
        value=10.0,
        step=1.0,
    )

    clearance = st.number_input(
        "Espacement entre pièces (mm)",
        min_value=0.0,
        value=5.0,
        step=1.0,
    )

    tolerance = st.number_input(
        "Précision des courbes DXF (mm)",
        min_value=0.1,
        value=1.0,
        step=0.1,
        help="Plus la valeur est petite, plus les arcs sont précis mais le calcul est lourd.",
    )

    rotation_options = st.multiselect(
        "Rotations autorisées",
        options=[0, 90, 180, 270],
        default=[0, 90, 180, 270],
    )

    st.warning(
        "Ce moteur est une version MVP heuristique. "
        "Vérifie toujours le DXF exporté avant une découpe réelle."
    )

st.subheader("1. Charger les fichiers")

left, right = st.columns(2)

with left:
    dxf_zip = st.file_uploader(
        "ZIP contenant les fichiers DXF",
        type=["zip"],
    )

with right:
    nomenclature_file = st.file_uploader(
        "Nomenclature Excel ou CSV",
        type=["xlsx", "csv"],
    )

st.subheader("2. Lancer l'analyse et l'optimisation")

run_button = st.button(
    "Analyser les DXF et optimiser les tôles",
    type="primary",
    use_container_width=True,
)

if run_button:
    if dxf_zip is None or nomenclature_file is None:
        st.warning("Charge le ZIP des DXF et la nomenclature.")
        st.stop()

    if not rotation_options:
        st.warning("Choisis au moins une rotation.")
        st.stop()

    try:
        with st.spinner("Lecture de la nomenclature..."):
            nomenclature = read_nomenclature(nomenclature_file)

        with st.spinner("Lecture et contrôle des DXF..."):
            dxf_files = read_dxf_zip(dxf_zip)
            pieces, missing, unused = build_pieces(
                nomenclature,
                dxf_files,
                tolerance,
            )

        if missing:
            st.error(
                "DXF manquants pour les repères : "
                + ", ".join(missing)
            )

        if unused:
            st.warning(
                "DXF présents mais non utilisés : "
                + ", ".join(unused)
            )

        if not pieces:
            st.error("Aucune pièce ne peut être optimisée.")
            st.stop()

        groups = defaultdict(list)
        for piece in pieces:
            groups[(piece.material, piece.thickness)].append(piece)

        all_placements: list[Placement] = []
        group_reports = []
        export_files = io.BytesIO()

        with zipfile.ZipFile(export_files, "w", zipfile.ZIP_DEFLATED) as master_zip:
            global_sheet_offset = 0

            for (material, thickness), group_pieces in groups.items():
                with st.spinner(
                    f"Optimisation : {material} — épaisseur {thickness}..."
                ):
                    placements, sheet_count, unplaced = nest_pieces(
                        group_pieces,
                        sheet_width,
                        sheet_height,
                        margin,
                        clearance,
                        sorted(rotation_options),
                    )

                for placement in placements:
                    placement.sheet_index += global_sheet_offset

                all_placements.extend(placements)

                group_reports.append(
                    {
                        "Matière": material,
                        "Épaisseur": thickness,
                        "Tôles utilisées": sheet_count,
                        "Pièces placées": len(placements),
                        "Pièces non placées": len(unplaced),
                    }
                )

                if unplaced:
                    st.error(
                        f"Pièces trop grandes ou non placées pour "
                        f"{material} / {thickness} : "
                        + ", ".join(unplaced)
                    )

                group_zip = export_sheets_zip(
                    [
                        Placement(
                            reference=p.reference,
                            source_name=p.source_name,
                            sheet_index=p.sheet_index - global_sheet_offset,
                            rotation=p.rotation,
                            polygon=p.polygon,
                            original_polygon=p.original_polygon,
                            copy_index=p.copy_index,
                            thickness=p.thickness,
                            material=p.material,
                        )
                        for p in placements
                    ],
                    sheet_count,
                    sheet_width,
                    sheet_height,
                )

                safe_material = re.sub(r"[^A-Za-z0-9_-]+", "_", material)
                safe_thickness = re.sub(r"[^A-Za-z0-9_-]+", "_", thickness)

                with zipfile.ZipFile(io.BytesIO(group_zip)) as inner_zip:
                    for inner_name in inner_zip.namelist():
                        master_zip.writestr(
                            f"{safe_material}_{safe_thickness}/{inner_name}",
                            inner_zip.read(inner_name),
                        )

                global_sheet_offset += sheet_count

        total_sheets = global_sheet_offset

        st.success(
            f"Optimisation terminée : {len(all_placements)} pièce(s) "
            f"placée(s) sur {total_sheets} tôle(s)."
        )

        st.subheader("3. Résumé par matière et épaisseur")
        st.dataframe(
            pd.DataFrame(group_reports),
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("4. Résultats détaillés")
        details = placement_table(all_placements)
        st.dataframe(
            details,
            use_container_width=True,
            hide_index=True,
        )

        st.subheader("5. Aperçu des tôles")
        if total_sheets > 0:
            selected_sheet = st.selectbox(
                "Choisir une tôle",
                options=list(range(1, total_sheets + 1)),
            )

            figure = plot_sheet(
                all_placements,
                selected_sheet - 1,
                sheet_width,
                sheet_height,
            )
            st.pyplot(figure, clear_figure=True)

        st.subheader("6. Téléchargements")

        details_csv = details.to_csv(
            index=False,
            sep=";",
        ).encode("utf-8-sig")

        download_left, download_right = st.columns(2)

        with download_left:
            st.download_button(
                "Télécharger le rapport CSV",
                data=details_csv,
                file_name="rapport_optitole.csv",
                mime="text/csv",
                use_container_width=True,
            )

        with download_right:
            st.download_button(
                "Télécharger les tôles DXF",
                data=export_files.getvalue(),
                file_name="optitole_resultats_dxf.zip",
                mime="application/zip",
                use_container_width=True,
            )

    except Exception as exc:
        st.exception(exc)

st.divider()
st.caption(
    "OptiTôle Pro MVP — Les contours DXF doivent être des polylignes fermées. "
    "Le résultat doit être contrôlé avant toute utilisation en production."
)
