import csv
import io
import math
import random
import re
import tempfile
import time
import unicodedata
import zipfile
from collections import defaultdict
from dataclasses import dataclass
from numbers import Number
from pathlib import Path
from typing import Iterable

import ezdxf
from ezdxf import edgeminer, edgesmith
import matplotlib.pyplot as plt
import pandas as pd
import streamlit as st
from ezdxf.path import from_hatch, make_path
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

def flatten_path(path, tolerance: float) -> list[tuple[float, float]]:
    try:
        return [
            (float(vertex.x), float(vertex.y))
            for vertex in path.flattening(max(tolerance, 0.05))
        ]
    except Exception:
        return []


def flatten_entity(entity, tolerance: float) -> list[tuple[float, float]]:
    try:
        return flatten_path(make_path(entity), tolerance)
    except Exception:
        return []


def clean_ring(points: list[tuple[float, float]]) -> list[tuple[float, float]]:
    cleaned: list[tuple[float, float]] = []

    for x, y in points:
        point = (round(float(x), 6), round(float(y), 6))
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


def polygon_from_ring_points(
    points: list[tuple[float, float]],
    minimum_area: float = 0.01,
) -> Polygon | None:
    points = clean_ring(points)

    if len(points) < 3:
        return None

    try:
        polygon = Polygon(LinearRing(points))
    except Exception:
        return None

    if polygon.is_empty or polygon.area <= minimum_area:
        return None

    if not polygon.is_valid:
        polygon = make_valid(polygon)

    if polygon.geom_type == "MultiPolygon":
        polygon = max(polygon.geoms, key=lambda part: part.area)

    if polygon.geom_type != "Polygon" or polygon.area <= minimum_area:
        return None

    return polygon


def deduplicate_polygons(polygons: list[Polygon]) -> list[Polygon]:
    result: list[Polygon] = []

    for polygon in sorted(polygons, key=lambda item: item.area, reverse=True):
        duplicate = False

        for existing in result:
            area_scale = max(existing.area, polygon.area, 1.0)

            if abs(existing.area - polygon.area) / area_scale > 1e-5:
                continue

            if existing.hausdorff_distance(polygon) <= 0.01:
                duplicate = True
                break

        if not duplicate:
            result.append(polygon)

    return result


def iter_geometry_entities(layout, maximum_depth: int = 8):
    """
    Parcourt les entités du modelspace et développe récursivement les INSERT.

    Advance Steel peut stocker certains contours ou perçages dans des blocs.
    INSERT.virtual_entities() renvoie les entités déjà transformées en coordonnées
    du dessin, sans modifier le fichier source.
    """
    stack = [(entity, 0) for entity in layout]

    while stack:
        entity, depth = stack.pop(0)

        if entity.dxftype() == "INSERT" and depth < maximum_depth:
            try:
                virtual = list(entity.virtual_entities())
                stack[0:0] = [(child, depth + 1) for child in virtual]
            except Exception:
                continue
        else:
            yield entity


def extract_direct_closed_loops(entities, tolerance: float) -> list[Polygon]:
    polygons: list[Polygon] = []

    for entity in entities:
        entity_type = entity.dxftype()

        if entity_type in {"HATCH", "MPOLYGON"}:
            try:
                for path in from_hatch(entity):
                    polygon = polygon_from_ring_points(
                        flatten_path(path, tolerance)
                    )
                    if polygon is not None:
                        polygons.append(polygon)
            except Exception:
                continue

        elif entity_is_closed(entity):
            polygon = polygon_from_ring_points(
                flatten_entity(entity, tolerance)
            )
            if polygon is not None:
                polygons.append(polygon)

    return polygons


def extract_connected_edge_loops(
    entities,
    tolerance: float,
    gap_tolerance: float,
) -> list[Polygon]:
    """Reconstruit les contours composés de LINE, ARC et courbes ouvertes."""
    polygons: list[Polygon] = []

    try:
        open_entities = list(edgesmith.filter_open_edges(entities))
        edges = list(
            edgesmith.edges_from_entities_2d(
                open_entities,
                gap_tol=gap_tolerance,
            )
        )

        if len(edges) < 2:
            return polygons

        deposit = edgeminer.Deposit(edges, gap_tol=gap_tolerance)
        loops = edgeminer.find_all_loops(deposit, timeout=8.0)

        for loop in loops:
            try:
                path = edgesmith.path2d_from_chain(loop)
                polygon = polygon_from_ring_points(
                    flatten_path(path, tolerance)
                )
                if polygon is not None:
                    polygons.append(polygon)
            except Exception:
                continue

    except Exception:
        # Les contours fermés directs restent exploitables même si un DXF
        # contient un réseau de lignes trop complexe pour EdgeMiner.
        return polygons

    return polygons


def build_plate_polygon(loop_polygons: list[Polygon], source_name: str) -> Polygon:
    loops = deduplicate_polygons(loop_polygons)

    if not loops:
        raise ValueError(
            f"Aucun contour fermé exploitable dans {source_name}."
        )

    loops.sort(key=lambda polygon: polygon.area, reverse=True)
    outer = loops[0]
    inside_loops = []

    for candidate in loops[1:]:
        point = candidate.representative_point()
        if outer.covers(point):
            inside_loops.append(candidate)

    # Les trous directs sont les boucles dont le plus petit contenant est
    # le contour extérieur. Les éventuels îlots imbriqués ne deviennent pas
    # de faux trous supplémentaires.
    direct_holes: list[Polygon] = []

    for candidate in inside_loops:
        candidate_point = candidate.representative_point()
        has_inner_parent = any(
            other.area > candidate.area
            and other.covers(candidate_point)
            for other in inside_loops
            if other is not candidate
        )

        if not has_inner_parent:
            direct_holes.append(candidate)

    plate = Polygon(
        outer.exterior.coords[:],
        [hole.exterior.coords[:] for hole in direct_holes],
    )

    if not plate.is_valid:
        plate = make_valid(plate)

    if plate.geom_type == "MultiPolygon":
        plate = max(plate.geoms, key=lambda part: part.area)

    if plate.geom_type != "Polygon" or plate.is_empty:
        raise ValueError(
            f"La géométrie de {source_name} est invalide après reconstruction."
        )

    min_x, min_y, _, _ = plate.bounds
    return affinity.translate(plate, xoff=-min_x, yoff=-min_y)


def polygons_from_dxf_bytes(
    data: bytes,
    source_name: str,
    tolerance: float,
) -> list[Polygon]:
    with tempfile.NamedTemporaryFile(suffix=".dxf", delete=False) as tmp:
        tmp.write(data)
        temp_name = tmp.name

    try:
        document = ezdxf.readfile(temp_name)
        modelspace = document.modelspace()
        geometry_entities = list(iter_geometry_entities(modelspace))

        gap_tolerance = max(0.02, min(0.5, tolerance * 0.2))

        loop_polygons = extract_direct_closed_loops(
            geometry_entities,
            tolerance,
        )
        loop_polygons.extend(
            extract_connected_edge_loops(
                geometry_entities,
                tolerance,
                gap_tolerance,
            )
        )

        return [build_plate_polygon(loop_polygons, source_name)]

    except ValueError:
        raise
    except Exception as exc:
        raise ValueError(f"DXF illisible : {source_name}") from exc
    finally:
        try:
            Path(temp_name).unlink(missing_ok=True)
        except Exception:
            pass


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
    rotated = affinity.rotate(
        polygon,
        angle,
        origin=(0, 0),
        use_radians=False,
    )
    min_x, min_y, _, _ = rotated.bounds
    return affinity.translate(rotated, xoff=-min_x, yoff=-min_y)


def sampled_coordinates(ring, maximum: int = 12) -> list[tuple[float, float]]:
    coordinates = list(ring.coords)

    if len(coordinates) <= maximum:
        return [(float(x), float(y)) for x, y in coordinates]

    step = max(1, len(coordinates) // maximum)
    return [
        (float(coordinates[index][0]), float(coordinates[index][1]))
        for index in range(0, len(coordinates), step)
    ][:maximum]


def candidate_positions(
    rotated_piece: Polygon,
    placed_polygons: list[Polygon],
    margin: float,
    clearance: float,
    sheet_width: float,
    sheet_height: float,
    allow_hole_nesting: bool,
    quality_level: int,
    max_candidates: int,
) -> list[tuple[float, float]]:
    """
    Génère des positions bas-gauche, autour des boîtes englobantes et par
    alignement de sommets. L'alignement de sommets permet aux formes inclinées
    et concaves de mieux s'emboîter qu'avec une simple mise en rangées.
    """
    width = rotated_piece.bounds[2] - rotated_piece.bounds[0]
    height = rotated_piece.bounds[3] - rotated_piece.bounds[1]
    candidates = {(round(margin, 4), round(margin, 4))}

    moving_vertices = sampled_coordinates(
        rotated_piece.exterior,
        maximum=8 if quality_level <= 1 else 14,
    )

    def add_candidate(x: float, y: float) -> None:
        if (
            x >= margin - 1e-6
            and y >= margin - 1e-6
            and x + width <= sheet_width - margin + 1e-6
            and y + height <= sheet_height - margin + 1e-6
        ):
            candidates.add((round(x, 4), round(y, 4)))

    for polygon in placed_polygons:
        min_x, min_y, max_x, max_y = polygon.bounds
        right_x = max_x + clearance
        top_y = max_y + clearance
        left_x = min_x - clearance - width
        bottom_y = min_y - clearance - height

        for x, y in (
            (right_x, min_y),
            (right_x, margin),
            (min_x, top_y),
            (margin, top_y),
            (right_x, top_y),
            (left_x, min_y),
            (min_x, bottom_y),
        ):
            add_candidate(x, y)

        fixed_vertices = sampled_coordinates(
            polygon.exterior,
            maximum=8 if quality_level <= 1 else 16,
        )

        # Alignement de sommets du contour mobile avec ceux des pièces fixes.
        # Cela fournit des positions d'emboîtement proches d'un placement
        # true-shape sans calculer un NFP complet.
        if quality_level >= 1:
            for fixed_x, fixed_y in fixed_vertices:
                for moving_x, moving_y in moving_vertices:
                    add_candidate(
                        fixed_x + clearance - moving_x,
                        fixed_y - moving_y,
                    )
                    add_candidate(
                        fixed_x - clearance - moving_x,
                        fixed_y - moving_y,
                    )
                    add_candidate(
                        fixed_x - moving_x,
                        fixed_y + clearance - moving_y,
                    )
                    add_candidate(
                        fixed_x - moving_x,
                        fixed_y - clearance - moving_y,
                    )
                    if len(candidates) >= max_candidates * 3:
                        break
                if len(candidates) >= max_candidates * 3:
                    break

        if allow_hole_nesting:
            for interior in polygon.interiors:
                hole = Polygon(interior)
                hole_min_x, hole_min_y, hole_max_x, hole_max_y = hole.bounds
                hole_width = hole_max_x - hole_min_x
                hole_height = hole_max_y - hole_min_y

                if (
                    hole_width + 1e-6 >= width + 2 * clearance
                    and hole_height + 1e-6 >= height + 2 * clearance
                ):
                    for x, y in (
                        (hole_min_x + clearance, hole_min_y + clearance),
                        (hole_max_x - clearance - width, hole_min_y + clearance),
                        (hole_min_x + clearance, hole_max_y - clearance - height),
                        (hole.centroid.x - width / 2.0, hole.centroid.y - height / 2.0),
                    ):
                        add_candidate(x, y)

                    if quality_level >= 2:
                        hole_vertices = sampled_coordinates(interior, maximum=12)
                        for fixed_x, fixed_y in hole_vertices:
                            for moving_x, moving_y in moving_vertices:
                                add_candidate(fixed_x - moving_x, fixed_y - moving_y)
                                if len(candidates) >= max_candidates * 3:
                                    break
                            if len(candidates) >= max_candidates * 3:
                                break

        if len(candidates) >= max_candidates * 3:
            break

    return sorted(
        candidates,
        key=lambda position: (position[1], position[0]),
    )[:max_candidates]


def bounds_are_close(
    first_bounds: tuple[float, float, float, float],
    second_bounds: tuple[float, float, float, float],
    clearance: float,
) -> bool:
    first_min_x, first_min_y, first_max_x, first_max_y = first_bounds
    second_min_x, second_min_y, second_max_x, second_max_y = second_bounds

    if first_max_x + clearance <= second_min_x:
        return False
    if second_max_x + clearance <= first_min_x:
        return False
    if first_max_y + clearance <= second_min_y:
        return False
    if second_max_y + clearance <= first_min_y:
        return False

    return True


def fits_on_sheet(
    polygon: Polygon,
    placed_polygons: list[Polygon],
    sheet_inner: Polygon,
    clearance: float,
) -> bool:
    if not sheet_inner.covers(polygon):
        return False

    candidate_bounds = polygon.bounds

    for other in placed_polygons:
        if not bounds_are_close(
            candidate_bounds,
            other.bounds,
            clearance,
        ):
            continue

        # Les trous sont respectés par Shapely : une pièce entièrement
        # placée dans un trou ne croise pas la matière de la grande pièce.
        if polygon.intersects(other):
            return False

        if clearance > 0 and polygon.distance(other) < clearance - 1e-7:
            return False

    return True


def compact_candidate(
    polygon: Polygon,
    placed_polygons: list[Polygon],
    sheet_inner: Polygon,
    clearance: float,
    minimum_step: float = 0.5,
) -> Polygon:
    """Fait glisser une pièce vers la gauche puis vers le bas."""
    result = polygon

    for _ in range(2):
        for axis in ("x", "y"):
            min_x, min_y, _, _ = result.bounds
            available = (
                min_x - sheet_inner.bounds[0]
                if axis == "x"
                else min_y - sheet_inner.bounds[1]
            )
            step = max(0.0, available)

            while step >= minimum_step:
                moved = affinity.translate(
                    result,
                    xoff=-step if axis == "x" else 0,
                    yoff=-step if axis == "y" else 0,
                )
                if fits_on_sheet(
                    moved,
                    placed_polygons,
                    sheet_inner,
                    clearance,
                ):
                    result = moved
                else:
                    step /= 2.0

    return result


def placement_score(
    polygon: Polygon,
    current_max_x: float,
    current_max_y: float,
) -> tuple:
    min_x, min_y, max_x, max_y = polygon.bounds
    new_max_x = max(current_max_x, max_x)
    new_max_y = max(current_max_y, max_y)
    used_rectangle = new_max_x * new_max_y

    return (
        round(new_max_y, 4),
        round(used_rectangle, 2),
        round(new_max_x, 4),
        round(min_y, 4),
        round(min_x, 4),
    )


def expand_piece_copies(
    pieces: list[DxfPiece],
) -> list[tuple[DxfPiece, int]]:
    expanded: list[tuple[DxfPiece, int]] = []

    for piece in pieces:
        for copy_index in range(1, piece.quantity + 1):
            expanded.append((piece, copy_index))

    return expanded


def order_piece_copies(
    expanded: list[tuple[DxfPiece, int]],
    attempt_index: int,
) -> list[tuple[DxfPiece, int]]:
    ordered = list(expanded)
    strategy = attempt_index % 6

    if strategy == 0:
        ordered.sort(key=lambda item: item[0].polygon.area, reverse=True)
    elif strategy == 1:
        ordered.sort(
            key=lambda item: (
                (
                    item[0].polygon.bounds[2]
                    - item[0].polygon.bounds[0]
                )
                * (
                    item[0].polygon.bounds[3]
                    - item[0].polygon.bounds[1]
                )
            ),
            reverse=True,
        )
    elif strategy == 2:
        ordered.sort(
            key=lambda item: max(
                item[0].polygon.bounds[2]
                - item[0].polygon.bounds[0],
                item[0].polygon.bounds[3]
                - item[0].polygon.bounds[1],
            ),
            reverse=True,
        )
    elif strategy == 3:
        ordered.sort(key=lambda item: item[0].polygon.length, reverse=True)
    elif strategy == 4:
        ordered.sort(
            key=lambda item: (
                sum(
                    Polygon(ring).area
                    for ring in item[0].polygon.interiors
                ),
                item[0].polygon.area,
            ),
            reverse=True,
        )
    else:
        random.Random(1000 + attempt_index).shuffle(ordered)

    return ordered


def compactness_score(
    sheets: list[list[Placement]],
    margin: float,
) -> tuple:
    used_heights = []
    used_areas = []

    for sheet in sheets:
        if not sheet:
            continue

        max_x = max(item.polygon.bounds[2] for item in sheet)
        max_y = max(item.polygon.bounds[3] for item in sheet)
        used_heights.append(max_y - margin)
        used_areas.append(
            (max_x - margin) * (max_y - margin)
        )

    return (
        len(sheets),
        round(sum(used_heights), 2),
        round(sum(used_areas), 2),
    )


def fallback_shelf_nest(
    pieces: list[DxfPiece],
    sheet_width: float,
    sheet_height: float,
    margin: float,
    clearance: float,
    rotations: list[int],
    rotation_cache: dict[tuple[str, int], Polygon],
) -> tuple[list[Placement], list[list[Placement]], list[str]]:
    """
    Placement de secours très rapide par rangées.

    Cette solution est calculée avant les essais avancés. Elle garantit
    qu'un résultat complet reste disponible même si la limite de temps
    est atteinte pendant l'amélioration.
    """
    inner_right = sheet_width - margin
    inner_top = sheet_height - margin
    expanded = expand_piece_copies(pieces)

    expanded.sort(
        key=lambda item: (
            max(
                item[0].polygon.bounds[2]
                - item[0].polygon.bounds[0],
                item[0].polygon.bounds[3]
                - item[0].polygon.bounds[1],
            ),
            item[0].polygon.area,
        ),
        reverse=True,
    )

    simple_angles = [
        angle
        for angle in (0, 90)
        if angle in rotations
    ]

    if not simple_angles:
        simple_angles = [rotations[0]]

    sheets_data: list[dict] = []
    unplaced: list[str] = []

    for piece, copy_index in expanded:
        rotated_versions = []

        for angle in simple_angles:
            cache_key = (piece.reference_key, angle)

            if cache_key not in rotation_cache:
                rotation_cache[cache_key] = rotated_at_origin(
                    piece.polygon,
                    angle,
                )

            polygon = rotation_cache[cache_key]
            width = polygon.bounds[2] - polygon.bounds[0]
            height = polygon.bounds[3] - polygon.bounds[1]
            rotated_versions.append(
                (angle, polygon, width, height)
            )

        best_option = None

        for sheet_index, sheet_data in enumerate(sheets_data):
            shelves = sheet_data["shelves"]

            for shelf_index, shelf in enumerate(shelves):
                for angle, polygon, width, height in rotated_versions:
                    if (
                        height <= shelf["height"] + 1e-6
                        and shelf["x"] + width <= inner_right + 1e-6
                    ):
                        remaining = inner_right - (shelf["x"] + width)
                        score = (
                            0,
                            remaining,
                            shelf["y"],
                            sheet_index,
                        )

                        if best_option is None or score < best_option[0]:
                            best_option = (
                                score,
                                "existing_shelf",
                                sheet_index,
                                shelf_index,
                                angle,
                                polygon,
                                width,
                                height,
                            )

            new_shelf_y = (
                margin
                if not shelves
                else max(
                    shelf["y"] + shelf["height"] + clearance
                    for shelf in shelves
                )
            )

            for angle, polygon, width, height in rotated_versions:
                if (
                    margin + width <= inner_right + 1e-6
                    and new_shelf_y + height <= inner_top + 1e-6
                ):
                    score = (
                        1,
                        new_shelf_y + height,
                        width,
                        sheet_index,
                    )

                    if best_option is None or score < best_option[0]:
                        best_option = (
                            score,
                            "new_shelf",
                            sheet_index,
                            None,
                            angle,
                            polygon,
                            width,
                            height,
                        )

        if best_option is None:
            valid_new_sheet = []

            for angle, polygon, width, height in rotated_versions:
                if (
                    margin + width <= inner_right + 1e-6
                    and margin + height <= inner_top + 1e-6
                ):
                    valid_new_sheet.append(
                        (height, width, angle, polygon)
                    )

            if not valid_new_sheet:
                unplaced.append(
                    f"{piece.reference_display} - copie {copy_index}"
                )
                continue

            _, _, angle, polygon = min(valid_new_sheet)
            width = polygon.bounds[2] - polygon.bounds[0]
            height = polygon.bounds[3] - polygon.bounds[1]
            sheet_index = len(sheets_data)

            sheets_data.append(
                {
                    "placements": [],
                    "shelves": [
                        {
                            "y": margin,
                            "height": height,
                            "x": margin,
                        }
                    ],
                }
            )

            best_option = (
                (2, height, width, sheet_index),
                "existing_shelf",
                sheet_index,
                0,
                angle,
                polygon,
                width,
                height,
            )

        (
            _,
            option_type,
            sheet_index,
            shelf_index,
            angle,
            polygon,
            width,
            height,
        ) = best_option

        sheet_data = sheets_data[sheet_index]

        if option_type == "new_shelf":
            shelf_y = (
                margin
                if not sheet_data["shelves"]
                else max(
                    shelf["y"] + shelf["height"] + clearance
                    for shelf in sheet_data["shelves"]
                )
            )

            sheet_data["shelves"].append(
                {
                    "y": shelf_y,
                    "height": height,
                    "x": margin,
                }
            )
            shelf_index = len(sheet_data["shelves"]) - 1

        shelf = sheet_data["shelves"][shelf_index]
        x = shelf["x"]
        y = shelf["y"]
        placed_polygon = affinity.translate(
            polygon,
            xoff=x,
            yoff=y,
        )

        placement = Placement(
            reference=piece.reference_display,
            source_name=piece.source_name,
            sheet_index=sheet_index,
            rotation=angle,
            polygon=placed_polygon,
            original_polygon=piece.polygon,
            copy_index=copy_index,
            thickness=piece.thickness,
            material=piece.material,
        )

        sheet_data["placements"].append(placement)
        shelf["x"] = x + width + clearance
        shelf["height"] = max(shelf["height"], height)

    sheets = [
        sheet_data["placements"]
        for sheet_data in sheets_data
    ]
    placements = [
        item
        for sheet in sheets
        for item in sheet
    ]

    return placements, sheets, unplaced


def nest_one_order(
    ordered_pieces: list[tuple[DxfPiece, int]],
    sheet_width: float,
    sheet_height: float,
    margin: float,
    clearance: float,
    rotations: list[int],
    allow_hole_nesting: bool,
    quality_level: int,
    max_candidates: int,
    rotation_cache: dict[tuple[str, int], Polygon],
    deadline: float,
    item_progress_callback=None,
) -> tuple[
    list[Placement],
    list[list[Placement]],
    list[str],
    bool,
]:
    sheet_inner = box(
        margin,
        margin,
        sheet_width - margin,
        sheet_height - margin,
    )
    inner_width = sheet_width - 2 * margin
    inner_height = sheet_height - 2 * margin
    sheets: list[list[Placement]] = []
    unplaced: list[str] = []
    total = len(ordered_pieces)

    for item_index, (piece, copy_index) in enumerate(
        ordered_pieces,
        start=1,
    ):
        if time.perf_counter() >= deadline:
            return [], [], [], True

        rotated_versions: list[tuple[int, Polygon]] = []

        for angle in rotations:
            cache_key = (piece.reference_key, angle)

            if cache_key not in rotation_cache:
                rotation_cache[cache_key] = rotated_at_origin(
                    piece.polygon,
                    angle,
                )

            rotated_versions.append(
                (angle, rotation_cache[cache_key])
            )

        best_global = None

        for sheet_index, sheet_placements in enumerate(sheets):
            if time.perf_counter() >= deadline:
                return [], [], [], True

            current_polygons = [
                item.polygon
                for item in sheet_placements
            ]
            current_max_x = max(
                (
                    polygon.bounds[2]
                    for polygon in current_polygons
                ),
                default=margin,
            )
            current_max_y = max(
                (
                    polygon.bounds[3]
                    for polygon in current_polygons
                ),
                default=margin,
            )

            for angle, rotated in rotated_versions:
                if time.perf_counter() >= deadline:
                    return [], [], [], True

                width = rotated.bounds[2] - rotated.bounds[0]
                height = rotated.bounds[3] - rotated.bounds[1]

                if (
                    width > inner_width + 1e-6
                    or height > inner_height + 1e-6
                ):
                    continue

                positions = candidate_positions(
                    rotated,
                    current_polygons,
                    margin,
                    clearance,
                    sheet_width,
                    sheet_height,
                    allow_hole_nesting,
                    quality_level,
                    max_candidates,
                )

                for x, y in positions:
                    if time.perf_counter() >= deadline:
                        return [], [], [], True

                    candidate = affinity.translate(
                        rotated,
                        xoff=x,
                        yoff=y,
                    )

                    if not fits_on_sheet(
                        candidate,
                        current_polygons,
                        sheet_inner,
                        clearance,
                    ):
                        continue

                    if quality_level >= 2:
                        candidate = compact_candidate(
                            candidate,
                            current_polygons,
                            sheet_inner,
                            clearance,
                        )

                    local_score = placement_score(
                        candidate,
                        current_max_x,
                        current_max_y,
                    )
                    global_score = (
                        local_score[0],
                        local_score[1],
                        sheet_index,
                        *local_score[2:],
                    )

                    if (
                        best_global is None
                        or global_score < best_global[0]
                    ):
                        best_global = (
                            global_score,
                            sheet_index,
                            angle,
                            candidate,
                        )

        if best_global is not None:
            _, sheet_index, angle, candidate = best_global
            sheets[sheet_index].append(
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
        else:
            best_new_sheet = None

            for angle, rotated in rotated_versions:
                width = rotated.bounds[2] - rotated.bounds[0]
                height = rotated.bounds[3] - rotated.bounds[1]

                if (
                    width > inner_width + 1e-6
                    or height > inner_height + 1e-6
                ):
                    continue

                candidate = affinity.translate(
                    rotated,
                    xoff=margin,
                    yoff=margin,
                )

                if sheet_inner.covers(candidate):
                    score = (
                        round(candidate.bounds[3], 4),
                        round(candidate.bounds[2], 4),
                    )

                    if (
                        best_new_sheet is None
                        or score < best_new_sheet[0]
                    ):
                        best_new_sheet = (
                            score,
                            angle,
                            candidate,
                        )

            if best_new_sheet is None:
                unplaced.append(
                    f"{piece.reference_display} - copie {copy_index}"
                )
            else:
                _, angle, candidate = best_new_sheet
                sheet_index = len(sheets)
                sheets.append(
                    [
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
                    ]
                )

        if item_progress_callback is not None and total:
            item_progress_callback(item_index / total)

    placements = [
        item
        for sheet in sheets
        for item in sheet
    ]

    return placements, sheets, unplaced, False


def consolidate_last_sheets(
    sheets: list[list[Placement]],
    sheet_width: float,
    sheet_height: float,
    margin: float,
    clearance: float,
    rotations: list[int],
    allow_hole_nesting: bool,
    rotation_cache: dict[tuple[str, int], Polygon],
    deadline: float,
) -> list[list[Placement]]:
    """
    Essaie de vider les dernières tôles en réinsérant leurs pièces dans les
    tôles précédentes. Une modification n'est conservée que si toute la tôle
    source peut être supprimée.
    """
    sheet_inner = box(
        margin,
        margin,
        sheet_width - margin,
        sheet_height - margin,
    )

    changed = True
    while changed and len(sheets) > 1 and time.perf_counter() < deadline:
        changed = False
        source_index = len(sheets) - 1
        source_items = sorted(
            sheets[source_index],
            key=lambda item: item.original_polygon.area,
            reverse=True,
        )
        trial = [list(sheet) for sheet in sheets[:-1]]
        success = True

        for item in source_items:
            if time.perf_counter() >= deadline:
                success = False
                break

            best = None
            reference_key = normalize_reference(item.reference)

            for target_index, target in enumerate(trial):
                fixed = [placed.polygon for placed in target]
                current_max_x = max((p.bounds[2] for p in fixed), default=margin)
                current_max_y = max((p.bounds[3] for p in fixed), default=margin)

                for angle in rotations:
                    cache_key = (reference_key, angle)
                    if cache_key not in rotation_cache:
                        rotation_cache[cache_key] = rotated_at_origin(
                            item.original_polygon,
                            angle,
                        )
                    rotated = rotation_cache[cache_key]
                    positions = candidate_positions(
                        rotated,
                        fixed,
                        margin,
                        clearance,
                        sheet_width,
                        sheet_height,
                        allow_hole_nesting,
                        2,
                        90,
                    )

                    for x, y in positions:
                        candidate = affinity.translate(rotated, xoff=x, yoff=y)
                        if not fits_on_sheet(
                            candidate,
                            fixed,
                            sheet_inner,
                            clearance,
                        ):
                            continue
                        candidate = compact_candidate(
                            candidate,
                            fixed,
                            sheet_inner,
                            clearance,
                        )
                        score = (
                            target_index,
                            *placement_score(candidate, current_max_x, current_max_y),
                        )
                        if best is None or score < best[0]:
                            best = (score, target_index, angle, candidate)

            if best is None:
                success = False
                break

            _, target_index, angle, candidate = best
            trial[target_index].append(
                Placement(
                    reference=item.reference,
                    source_name=item.source_name,
                    sheet_index=target_index,
                    rotation=angle,
                    polygon=candidate,
                    original_polygon=item.original_polygon,
                    copy_index=item.copy_index,
                    thickness=item.thickness,
                    material=item.material,
                )
            )

        if success:
            sheets = trial
            changed = True

    for sheet_index, sheet in enumerate(sheets):
        for placement in sheet:
            placement.sheet_index = sheet_index

    return sheets


def nest_pieces(
    pieces: list[DxfPiece],
    sheet_width: float,
    sheet_height: float,
    margin: float,
    clearance: float,
    rotations: list[int],
    quality_mode: str = "Équilibré",
    allow_hole_nesting: bool = True,
    time_budget_seconds: float = 40.0,
    progress_callback=None,
) -> tuple[list[Placement], int, list[str]]:
    """
    Calcule d'abord une solution complète et rapide, puis tente de
    l'améliorer jusqu'à la limite de temps.

    Une solution est donc toujours renvoyée, même lorsque le calcul avancé
    est interrompu par le chronomètre.
    """
    if sheet_width <= 0 or sheet_height <= 0:
        raise ValueError(
            "Les dimensions de la tôle doivent être positives."
        )

    if margin < 0 or clearance < 0:
        raise ValueError(
            "La marge et l'espacement ne peuvent pas être négatifs."
        )

    if (
        sheet_width - 2 * margin <= 0
        or sheet_height - 2 * margin <= 0
    ):
        raise ValueError(
            "La marge est trop grande pour le format de tôle."
        )

    if not rotations:
        rotations = [0]

    expanded = expand_piece_copies(pieces)
    total_copies = len(expanded)
    rotation_cache: dict[tuple[str, int], Polygon] = {}

    # Réduction automatique de la complexité pour les gros projets.
    effective_rotations = sorted(set(rotations))

    if total_copies > 120:
        effective_rotations = [
            angle
            for angle in (0, 90)
            if angle in effective_rotations
        ] or [effective_rotations[0]]
    elif total_copies > 60 and len(effective_rotations) > 4:
        effective_rotations = [
            angle
            for angle in (0, 90, 180, 270)
            if angle in effective_rotations
        ] or effective_rotations[:4]

    quality_settings = {
        "Rapide": (1, 1, 35),
        "Équilibré": (3, 1, 65),
        "Approfondi": (6, 2, 100),
    }

    attempts, quality_level, max_candidates = quality_settings.get(
        quality_mode,
        (3, 1, 65),
    )

    if total_copies > 150:
        attempts = 1
        quality_level = 1
        max_candidates = 25
    elif total_copies > 80:
        attempts = min(attempts, 2)
        quality_level = 1
        max_candidates = min(max_candidates, 45)
    elif total_copies > 40:
        attempts = min(attempts, 3)
        max_candidates = min(max_candidates, 65)

    # 1. Solution de secours complète et immédiate.
    fallback_placements, fallback_sheets, fallback_unplaced = (
        fallback_shelf_nest(
            pieces,
            sheet_width,
            sheet_height,
            margin,
            clearance,
            effective_rotations,
            rotation_cache,
        )
    )

    best_result = (
        (
            len(fallback_unplaced),
            *compactness_score(fallback_sheets, margin),
        ),
        fallback_placements,
        fallback_sheets,
        fallback_unplaced,
    )

    if progress_callback is not None:
        progress_callback(0.08)

    # 2. Recherche d'améliorations dans un temps strictement limité.
    start_time = time.perf_counter()
    deadline = start_time + max(5.0, float(time_budget_seconds))

    for attempt_index in range(attempts):
        if time.perf_counter() >= deadline:
            break

        ordered = order_piece_copies(
            expanded,
            attempt_index,
        )

        def update_item_progress(item_fraction):
            if progress_callback is not None:
                fraction = (
                    0.08
                    + 0.90
                    * (
                        attempt_index + item_fraction
                    )
                    / max(1, attempts)
                )
                progress_callback(min(0.98, fraction))

        placements, sheets, unplaced, timed_out = nest_one_order(
            ordered,
            sheet_width,
            sheet_height,
            margin,
            clearance,
            effective_rotations,
            allow_hole_nesting,
            quality_level,
            max_candidates,
            rotation_cache,
            deadline,
            item_progress_callback=update_item_progress,
        )

        if timed_out:
            break

        score = (
            len(unplaced),
            *compactness_score(sheets, margin),
        )

        if score < best_result[0]:
            best_result = (
                score,
                placements,
                sheets,
                unplaced,
            )

        # Une seule tôle sans pièce manquante ne peut pas être améliorée
        # sur le critère principal du nombre de tôles.
        if not unplaced and len(sheets) <= 1:
            break

    if progress_callback is not None:
        progress_callback(1.0)

    _, placements, sheets, unplaced = best_result

    if not unplaced and len(sheets) > 1 and time.perf_counter() < deadline:
        sheets = consolidate_last_sheets(
            sheets,
            sheet_width,
            sheet_height,
            margin,
            clearance,
            effective_rotations,
            allow_hole_nesting,
            rotation_cache,
            deadline,
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
                "Trous": len(item.original_polygon.interiors),
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
    ax.set_title(
        f"Tôle {sheet_index + 1} — {sheet_width:g} × {sheet_height:g} mm"
    )
    ax.set_xlabel("X (mm)")
    ax.set_ylabel("Y (mm)")
    ax.grid(True, linewidth=0.3)

    sheet_items = [
        item for item in placements if item.sheet_index == sheet_index
    ]

    for item in sheet_items:
        x, y = item.polygon.exterior.xy
        filled = ax.fill(x, y, alpha=0.28)
        face_color = filled[0].get_facecolor()
        ax.plot(x, y, linewidth=1)

        # Les trous sont réellement évidés dans l'aperçu.
        for interior in item.polygon.interiors:
            hx, hy = interior.xy
            ax.fill(
                hx,
                hy,
                facecolor=ax.get_facecolor(),
                alpha=1.0,
            )
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

st.title("📐 OptiTôle Pro v9 Compact")
st.caption(
    "Nesting de tôles à partir de DXF Advance Steel et d'une nomenclature Excel/CSV. "
    "Version 9 : lecture des blocs DXF, diagnostic des trous et compactage amélioré."
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
        value=2.0,
        step=0.5,
        help=(
            "3 mm est conseillé pour un premier calcul rapide. "
            "Réduis ensuite à 1 mm pour le contrôle final."
        ),
    )

    rotation_step = st.selectbox(
        "Pas de rotation",
        options=[90, 45, 30, 15],
        index=0,
        help=(
            "90° est rapide. 45°, 30° ou 15° peuvent améliorer le rendement "
            "mais augmentent le temps de calcul."
        ),
    )

    rotation_options = list(range(0, 360, rotation_step))

    quality_mode = st.selectbox(
        "Qualité d'optimisation",
        options=["Rapide", "Équilibré", "Approfondi"],
        index=1,
        help=(
            "Le mode approfondi essaie davantage d'ordres de pièces. "
            "Il peut être sensiblement plus long."
        ),
    )

    time_budget_seconds = st.select_slider(
        "Temps maximum d'amélioration par projet (secondes)",
        options=[15, 30, 45, 60, 90],
        value=45,
        help=(
            "Une première solution complète est créée immédiatement. "
            "Le moteur utilise ensuite ce temps pour essayer de l'améliorer."
        ),
    )

    allow_hole_nesting = st.checkbox(
        "Autoriser les petites pièces dans les grands trous",
        value=True,
    )

    st.info(
        "La v9 développe aussi les blocs DXF, aligne les sommets des pièces "
        "et tente de vider les dernières tôles pour réduire les chutes."
    )

    st.warning(
        "Ce moteur reste un MVP heuristique. "
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

clear_results = st.button(
    "Effacer les anciens résultats",
    use_container_width=True,
)

if clear_results:
    st.session_state.pop("optitole_result", None)
    st.success("Les anciens résultats ont été effacés.")

if run_button:
    if dxf_zip is None or nomenclature_file is None:
        st.warning("Charge le ZIP des DXF et la nomenclature.")
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

        if not pieces:
            st.error("Aucune pièce ne peut être optimisée.")
            st.stop()

        dxf_diagnostic = pd.DataFrame(
            [
                {
                    "Repère": piece.reference_display,
                    "Fichier": piece.source_name,
                    "Trous détectés": len(piece.polygon.interiors),
                    "Surface nette (mm²)": round(piece.polygon.area, 1),
                }
                for piece in pieces
            ]
        )
        total_detected_holes = int(dxf_diagnostic["Trous détectés"].sum())

        if total_detected_holes == 0:
            st.warning(
                "Aucun contour de trou n'a été trouvé dans les DXF. "
                "Le fichier Advance Steel peut ne pas contenir les perçages, "
                "ou ceux-ci peuvent être exportés uniquement comme marques."
            )

        groups = defaultdict(list)

        for piece in pieces:
            groups[(piece.material, piece.thickness)].append(piece)

        all_placements: list[Placement] = []
        group_reports = []
        unplaced_messages = []
        export_files = io.BytesIO()
        total_project_copies = sum(
            piece.quantity for piece in pieces
        )

        with zipfile.ZipFile(
            export_files,
            "w",
            zipfile.ZIP_DEFLATED,
        ) as master_zip:
            global_sheet_offset = 0

            for (material, thickness), group_pieces in groups.items():
                total_group_copies = sum(
                    piece.quantity
                    for piece in group_pieces
                )

                st.write(
                    f"Optimisation : **{material} — épaisseur {thickness}** "
                    f"({total_group_copies} pièce(s))"
                )

                progress_bar = st.progress(0)

                def update_group_progress(value):
                    progress_bar.progress(
                        min(100, max(0, int(value * 100)))
                    )

                group_time_budget = max(
                    5.0,
                    float(time_budget_seconds)
                    * total_group_copies
                    / max(1, total_project_copies),
                )

                placements, sheet_count, unplaced = nest_pieces(
                    group_pieces,
                    sheet_width,
                    sheet_height,
                    margin,
                    clearance,
                    sorted(rotation_options),
                    quality_mode=quality_mode,
                    allow_hole_nesting=allow_hole_nesting,
                    time_budget_seconds=group_time_budget,
                    progress_callback=update_group_progress,
                )

                progress_bar.empty()

                for placement in placements:
                    placement.sheet_index += global_sheet_offset

                all_placements.extend(placements)

                group_reports.append(
                    {
                        "Matière": material,
                        "Épaisseur": thickness,
                        "Tôles utilisées": sheet_count,
                        "Pièces placées": len(placements),
                        "Trous détectés": sum(
                            len(piece.polygon.interiors) * piece.quantity
                            for piece in group_pieces
                        ),
                        "Pièces non placées": len(unplaced),
                    }
                )

                if unplaced:
                    unplaced_messages.append(
                        f"{material} / {thickness} : "
                        + ", ".join(unplaced)
                    )

                local_placements = [
                    Placement(
                        reference=p.reference,
                        source_name=p.source_name,
                        sheet_index=(
                            p.sheet_index
                            - global_sheet_offset
                        ),
                        rotation=p.rotation,
                        polygon=p.polygon,
                        original_polygon=p.original_polygon,
                        copy_index=p.copy_index,
                        thickness=p.thickness,
                        material=p.material,
                    )
                    for p in placements
                ]

                group_zip = export_sheets_zip(
                    local_placements,
                    sheet_count,
                    sheet_width,
                    sheet_height,
                )

                safe_material = re.sub(
                    r"[^A-Za-z0-9_-]+",
                    "_",
                    material,
                )

                safe_thickness = re.sub(
                    r"[^A-Za-z0-9_-]+",
                    "_",
                    thickness,
                )

                with zipfile.ZipFile(
                    io.BytesIO(group_zip)
                ) as inner_zip:
                    for inner_name in inner_zip.namelist():
                        master_zip.writestr(
                            (
                                f"{safe_material}_"
                                f"{safe_thickness}/"
                                f"{inner_name}"
                            ),
                            inner_zip.read(inner_name),
                        )

                global_sheet_offset += sheet_count

        total_sheets = global_sheet_offset
        details = placement_table(all_placements)

        details_csv = details.to_csv(
            index=False,
            sep=";",
        ).encode("utf-8-sig")

        st.session_state["optitole_result"] = {
            "placements": all_placements,
            "group_reports": group_reports,
            "total_sheets": total_sheets,
            "details": details,
            "details_csv": details_csv,
            "export_zip": export_files.getvalue(),
            "sheet_width": sheet_width,
            "sheet_height": sheet_height,
            "missing": missing,
            "unused": unused,
            "unplaced_messages": unplaced_messages,
            "time_budget_seconds": time_budget_seconds,
            "dxf_diagnostic": dxf_diagnostic,
        }

    except Exception as exc:
        st.exception(exc)


# Les résultats sont affichés en dehors du bouton.
# Ils restent donc visibles après un rafraîchissement ou un changement de tôle.
result = st.session_state.get("optitole_result")

if result is not None:
    missing = result["missing"]
    unused = result["unused"]
    unplaced_messages = result["unplaced_messages"]
    all_placements = result["placements"]
    group_reports = result["group_reports"]
    total_sheets = result["total_sheets"]
    details = result["details"]
    details_csv = result["details_csv"]
    export_zip = result["export_zip"]
    result_sheet_width = result["sheet_width"]
    result_sheet_height = result["sheet_height"]
    dxf_diagnostic = result.get("dxf_diagnostic")

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

    for message in unplaced_messages:
        st.error(
            "Pièces trop grandes ou non placées : "
            + message
        )

    st.success(
        f"Optimisation terminée : {len(all_placements)} pièce(s) "
        f"placée(s) sur {total_sheets} tôle(s)."
    )

    st.subheader("3. Diagnostic des contours DXF")
    if dxf_diagnostic is not None:
        st.dataframe(
            dxf_diagnostic,
            use_container_width=True,
            hide_index=True,
        )

    st.subheader("4. Résumé final")
    st.dataframe(
        pd.DataFrame(group_reports),
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("5. Résultats détaillés")
    st.dataframe(
        details,
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("6. Aperçu des tôles")

    if total_sheets > 0:
        selected_sheet = st.selectbox(
            "Choisir la tôle à afficher",
            options=list(range(1, total_sheets + 1)),
            key="selected_result_sheet",
        )

        figure = plot_sheet(
            all_placements,
            selected_sheet - 1,
            result_sheet_width,
            result_sheet_height,
        )

        st.pyplot(
            figure,
            clear_figure=True,
        )

    st.subheader("7. Télécharger les résultats")

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
            data=export_zip,
            file_name="optitole_resultats_dxf.zip",
            mime="application/zip",
            use_container_width=True,
        )
else:
    st.info(
        "Aucun résultat enregistré. "
        "Clique sur « Analyser les DXF et optimiser les tôles »."
    )

st.divider()
st.caption(
    "OptiTôle Pro v7 — Les contours fermés et les boucles LINE/ARC sont reconstruits. "
    "Le résultat doit être contrôlé avant toute utilisation en production."
)
