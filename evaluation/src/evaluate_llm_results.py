#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Évaluation automatique V4 et validation humaine des corrections produites par des LLM.

Entrée par défaut
-----------------
result/alignments.xlsx

Sortie par défaut
-----------------
result/evaluation_llm.xlsx

Fonctions principales
---------------------
1. Exactitude de la phrase corrigée après normalisation légère.
2. Précision, rappel et score F1 sur des éditions recalculées depuis les phrases.
3. Résultats globaux et résultats par catégorie P-Rel.
4. Préclassification automatique :
   - correction exacte ;
   - sous-correction ;
   - sur-correction ;
   - correction différente ;
   - sortie vide.
5. Conservation et agrégation des validations manuelles déjà saisies.
6. Tableaux et graphiques prêts à être utilisés dans le mémoire.

Validation manuelle
-------------------
Dans la feuille « Evaluation_detaillee », remplir la colonne
« Validation manuelle » avec l'une des valeurs suivantes :

- Correction acceptable
- Correction incorrecte
- Sous-correction
- Sur-correction
- Correction hors cible
- À discuter

La colonne « Commentaire humain » est libre.

Lors d'une nouvelle exécution, le programme tente de récupérer les validations
déjà présentes dans le fichier de sortie existant, à partir de la clé
(ID, Modèle).

Dépendances
-----------
pip install pandas openpyxl
"""

from __future__ import annotations

import argparse
import ast
import difflib
import json
import math
import re
import unicodedata
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Set, Tuple

import pandas as pd
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


MODELS: Dict[str, Dict[str, str]] = {
    "gpt_oss_20b": {
        "label": "GPT-OSS-20B",
        "output": "Sortie_gpt_oss_20b",
        "distance": "Distance_gpt_oss_20b",
        "edits": "Edits_gpt_oss_20b",
        "n_edits": "Nb_edits_gpt_oss_20b",
        "exact_edit": "Exact_edit_match_gold_gpt_oss_20b",
    },
    "deepseek_v4_flash": {
        "label": "DeepSeek-V4-Flash",
        "output": "Sortie_deepseek_v4_flash",
        "distance": "Distance_deepseek_v4_flash",
        "edits": "Edits_deepseek_v4_flash",
        "n_edits": "Nb_edits_deepseek_v4_flash",
        "exact_edit": "Exact_edit_match_gold_deepseek_v4_flash",
    },
    "qwen3_32b": {
        "label": "Qwen3-32B",
        "output": "Sortie_qwen3_32b",
        "distance": "Distance_qwen3_32b",
        "edits": "Edits_qwen3_32b",
        "n_edits": "Nb_edits_qwen3_32b",
        "exact_edit": "Exact_edit_match_gold_qwen3_32b",
    },
    "gemma3_27b": {
        "label": "Gemma3-27B",
        "output": "Sortie_gemma3_27b",
        "distance": "Distance_gemma3_27b",
        "edits": "Edits_gemma3_27b",
        "n_edits": "Nb_edits_gemma3_27b",
        "exact_edit": "Exact_edit_match_gold_gemma3_27b",
    },
}

BASE_COLUMNS = [
    "ID",
    "Phrase_erronée",
    "Phrase_gold",
    "Distance_gold",
    "Edits_gold",
    "Nb_edits_gold",
]

MANUAL_VALUES = [
    "Correction acceptable",
    "Correction incorrecte",
    "Sous-correction",
    "Sur-correction",
    "Correction hors cible",
    "À discuter",
]

SHEET_GLOBAL = "Synthese_globale"
SHEET_CATEGORY = "Resultats_par_categorie"
SHEET_MATRIX = "Matrice_categories"
SHEET_DETAIL = "Evaluation_detaillee"
SHEET_REVIEW = "Cas_a_verifier"
SHEET_MANUAL = "Statistiques_manuelles"
SHEET_CHARTS = "Graphiques_memoire"
SHEET_SOURCE = "Donnees_sources"
SHEET_METHOD = "Methodologie"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Évaluer les corrections de plusieurs LLM."
    )
    parser.add_argument(
        "--input",
        default="result/alignments.xlsx",
        help="Classeur contenant les alignements.",
    )
    parser.add_argument(
        "--output",
        default="result/evaluation_llm.xlsx",
        help="Classeur d'évaluation à produire.",
    )
    parser.add_argument(
        "--no-preserve-manual",
        action="store_true",
        help="Ne pas récupérer les validations manuelles du fichier de sortie existant.",
    )
    return parser.parse_args()


def as_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    return str(value).strip()


def as_number(value: Any) -> float:
    try:
        number = float(value)
        return 0.0 if math.isnan(number) else number
    except (TypeError, ValueError):
        return 0.0


def as_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return bool(value)
    return as_text(value).lower() in {
        "true", "vrai", "1", "yes", "oui"
    }


def safe_ratio(numerator: float, denominator: float) -> float:
    return numerator / denominator if denominator else 0.0


def normalize_sentence(value: Any) -> str:
    value = unicodedata.normalize("NFKC", as_text(value))
    replacements = {
        "’": "'",
        "‘": "'",
        "“": '"',
        "”": '"',
        "«": '"',
        "»": '"',
        "\u00a0": " ",
    }
    for old, new in replacements.items():
        value = value.replace(old, new)

    value = re.sub(r"\s+", " ", value)
    value = re.sub(r"\s+([,;:!?\.])", r"\1", value)
    value = re.sub(r"[.!?…]+$", "", value.strip())
    return value.casefold().strip()


def normalize_edit(value: Any) -> str:
    value = unicodedata.normalize("NFKC", as_text(value))
    value = value.replace("’", "'").replace("‘", "'")
    value = re.sub(r"\s+", " ", value)
    return value.casefold().strip(" \t\r\n,;")


def flatten_edit_object(obj: Any) -> List[str]:
    """Convertit plusieurs formats possibles d'éditions en signatures textuelles."""
    if obj is None:
        return []

    if isinstance(obj, dict):
        preferred_keys = [
            "type", "operation", "op", "tag",
            "source", "src", "old", "from",
            "target", "tgt", "new", "to",
            "start", "end", "position",
        ]
        parts = []
        used = set()
        for key in preferred_keys:
            if key in obj:
                parts.append(f"{key}={obj[key]}")
                used.add(key)
        for key in sorted(obj):
            if key not in used:
                parts.append(f"{key}={obj[key]}")
        return [" | ".join(parts)] if parts else []

    if isinstance(obj, (list, tuple, set)):
        result: List[str] = []
        for item in obj:
            result.extend(flatten_edit_object(item))
        return result

    return [as_text(obj)]




def normalize_for_edit_alignment(value: Any) -> str:
    """
    Normalisation utilisée uniquement pour l'alignement des éditions.

    Elle harmonise les apostrophes et les espaces, et ignore la ponctuation
    finale déjà ignorée par l'exactitude normalisée. Elle ne met pas le texte
    en minuscules ici : la casse est conservée dans les signatures lisibles,
    puis neutralisée par normalize_edit().
    """
    value = unicodedata.normalize("NFKC", as_text(value))
    replacements = {
        "’": "'",
        "‘": "'",
        "“": '"',
        "”": '"',
        "«": '"',
        "»": '"',
        "\u00a0": " ",
    }
    for old, new in replacements.items():
        value = value.replace(old, new)

    value = re.sub(r"\s+", " ", value).strip()
    value = re.sub(r"\s+([,;:!?\.])", r"\1", value)
    value = re.sub(r"[.!?…]+$", "", value).strip()
    return value


def tokenize_for_edits(value: Any) -> List[str]:
    """
    Tokenisation légère adaptée au français.

    Les mots avec apostrophe sont séparés de façon stable :
        "l'homme" -> ["l'", "homme"]
    La ponctuation interne devient un token indépendant.
    """
    text = normalize_for_edit_alignment(value)
    if not text:
        return []

    pattern = r"[A-Za-zÀ-ÖØ-öø-ÿ]+(?:['’])?|[0-9]+(?:[.,][0-9]+)?|[^\w\s]"
    return re.findall(pattern, text, flags=re.UNICODE)


def join_edit_tokens(tokens: Sequence[str]) -> str:
    """Recompose un segment lisible à partir de tokens."""
    if not tokens:
        return "∅"

    text = " ".join(tokens)
    text = re.sub(r"\s+([,;:!?\.])", r"\1", text)
    text = re.sub(r"([\'’])\s+", r"\1", text)
    return text.strip()


def compute_phrase_edits(source: Any, target: Any) -> List[str]:
    """
    Calcule les éditions entre une phrase source et une phrase cible.

    Chaque édition contient :
    - le type : suppression, insertion ou remplacement ;
    - la position dans la phrase source ;
    - le segment source ;
    - le segment cible.

    Les positions empêchent qu'une même forme répétée à deux endroits soit
    confondue. Une substitution est toujours une seule édition, même si
    l'algorithme interne l'exprime comme suppression + insertion.
    """
    source_tokens = tokenize_for_edits(source)
    target_tokens = tokenize_for_edits(target)

    matcher = difflib.SequenceMatcher(
        a=[token.casefold() for token in source_tokens],
        b=[token.casefold() for token in target_tokens],
        autojunk=False,
    )

    edits: List[str] = []

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            continue

        source_segment = join_edit_tokens(source_tokens[i1:i2])
        target_segment = join_edit_tokens(target_tokens[j1:j2])
        position = f"{i1}:{i2}"

        if tag == "delete":
            signature = (
                f"suppression@{position}:"
                f"{source_segment}"
            )
        elif tag == "insert":
            signature = (
                f"insertion@{position}:"
                f"{target_segment}"
            )
        else:
            signature = (
                f"remplacement@{position}:"
                f"{source_segment}→{target_segment}"
            )

        edits.append(normalize_edit(signature))

    return edits


def parse_edits(value: Any) -> List[str]:
    """
    Compatibilité avec les anciennes colonnes Edits_*.

    Cette fonction reste disponible pour l'affichage et les contrôles, mais
    les métriques principales V4 ne reposent plus sur ces cellules. Elles sont
    recalculées directement à partir des phrases avec compute_phrase_edits().
    """
    raw = as_text(value)
    if not raw or raw.lower() in {"nan", "none", "[]", "{}"}:
        return []

    parsed: Any = None
    for loader in (json.loads, ast.literal_eval):
        try:
            parsed = loader(raw)
            break
        except Exception:
            parsed = None

    if parsed is not None and not isinstance(parsed, str):
        candidates = flatten_edit_object(parsed)
    elif "\n" in raw:
        candidates = raw.splitlines()
    elif "||" in raw:
        candidates = raw.split("||")
    elif re.search(r"\s+\|\s+", raw):
        candidates = re.split(r"\s+\|\s+", raw)
    else:
        parts = [part for part in raw.split(";") if part.strip()]
        candidates = parts if len(parts) > 1 else [raw]

    normalized = [normalize_edit(item) for item in candidates]
    return [item for item in normalized if item]



def multiset_overlap(gold: Sequence[str], predicted: Sequence[str]) -> int:
    """Intersection de deux listes en tenant compte des doublons."""
    remaining = list(gold)
    matches = 0
    for item in predicted:
        if item in remaining:
            matches += 1
            remaining.remove(item)
    return matches


def edit_metrics(gold_edits: Sequence[str], predicted_edits: Sequence[str]) -> Dict[str, float]:
    tp = multiset_overlap(gold_edits, predicted_edits)
    fp = max(len(predicted_edits) - tp, 0)
    fn = max(len(gold_edits) - tp, 0)
    precision = safe_ratio(tp, tp + fp)
    recall = safe_ratio(tp, tp + fn)
    f1 = safe_ratio(2 * precision * recall, precision + recall)

    return {
        "TP_edition": tp,
        "FP_edition": fp,
        "FN_edition": fn,
        "Precision_edition": precision,
        "Rappel_edition": recall,
        "F1_edition": f1,
    }


def category_from_row(row: pd.Series) -> str:
    combined = " | ".join(
        as_text(row.get(column, ""))
        for column in ("Habileté", "Corrections", "Source")
    )
    match = re.search(
        r"\bP[-_ ]?Rel(?:0(?:_2|_3)?|1(?:_2)?)\b",
        combined,
        flags=re.IGNORECASE,
    )
    if not match:
        return "Non renseignée"

    return re.sub(
        r"^P[-_ ]?Rel",
        "P-Rel",
        match.group(0),
        flags=re.IGNORECASE,
    )


def validate_columns(df: pd.DataFrame) -> None:
    required = list(BASE_COLUMNS)

    for config in MODELS.values():
        required.extend([
            config["output"],
            config["distance"],
            config["edits"],
            config["n_edits"],
            config["exact_edit"],
        ])

    missing = [column for column in required if column not in df.columns]
    if missing:
        raise ValueError(
            "Colonnes manquantes dans le fichier d'entrée :\n"
            + "\n".join(f"- {column}" for column in missing)
        )


def automatic_diagnosis(
    source: str,
    output: str,
    exact_sentence: bool,
    gold_count: float,
    model_count: float,
) -> str:
    if not output:
        return "Sortie vide"
    if exact_sentence:
        return "Correction exacte"
    if normalize_sentence(output) == normalize_sentence(source):
        return "Sous-correction : aucune modification"
    if model_count < gold_count:
        return "Sous-correction probable"
    if model_count > gold_count:
        return "Sur-correction probable"
    return "Correction différente à vérifier"


def load_manual_annotations(path: Path) -> Dict[Tuple[str, str], Dict[str, str]]:
    """
    Récupère les validations humaines du classeur de sortie précédent.
    La clé est (ID, Modèle).
    """
    if not path.exists():
        return {}

    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return {}

    if SHEET_DETAIL not in workbook.sheetnames:
        return {}

    sheet = workbook[SHEET_DETAIL]
    headers = {
        as_text(cell.value): index
        for index, cell in enumerate(next(sheet.iter_rows()), start=1)
    }

    required = ["ID", "Modèle", "Validation manuelle", "Commentaire humain"]
    if any(column not in headers for column in required):
        return {}

    annotations: Dict[Tuple[str, str], Dict[str, str]] = {}
    for row in sheet.iter_rows(min_row=2, values_only=True):
        identifier = as_text(row[headers["ID"] - 1])
        model = as_text(row[headers["Modèle"] - 1])
        validation = as_text(row[headers["Validation manuelle"] - 1])
        comment = as_text(row[headers["Commentaire humain"] - 1])

        if identifier and model and (validation or comment):
            annotations[(identifier, model)] = {
                "Validation manuelle": validation,
                "Commentaire humain": comment,
            }

    workbook.close()
    return annotations



def build_detail(
    df: pd.DataFrame,
    annotations: Dict[Tuple[str, str], Dict[str, str]],
) -> pd.DataFrame:
    """
    Construit l'évaluation détaillée.

    V4 : les éditions de référence et du modèle sont recalculées à partir des
    phrases. Les anciennes colonnes Edits_* sont conservées uniquement comme
    traces d'alignement, afin que les métriques ne dépendent plus d'un format
    ou d'une segmentation externe instable.
    """
    rows: List[Dict[str, Any]] = []

    for _, row in df.iterrows():
        identifier = as_text(row["ID"])
        source = as_text(row["Phrase_erronée"])
        gold = as_text(row["Phrase_gold"])
        gold_normalized = normalize_sentence(gold)
        gold_edit_list = compute_phrase_edits(source, gold)
        gold_count = len(gold_edit_list)
        category = category_from_row(row)

        for config in MODELS.values():
            model = config["label"]
            output = as_text(row[config["output"]])
            output_normalized = normalize_sentence(output)
            exact_sentence = bool(
                gold_normalized and output_normalized == gold_normalized
            )

            predicted_edit_list = compute_phrase_edits(source, output)
            model_count = len(predicted_edit_list)
            metrics = edit_metrics(gold_edit_list, predicted_edit_list)
            exact_edit_recomputed = (
                sorted(gold_edit_list) == sorted(predicted_edit_list)
            )
            annotation = annotations.get((identifier, model), {})

            rows.append({
                "ID": identifier,
                "Source": as_text(row.get("Source", "")),
                "Catégorie": category,
                "Habileté": as_text(row.get("Habileté", "")),
                "Corrections": as_text(row.get("Corrections", "")),
                "Modèle": model,
                "Phrase erronée": source,
                "Phrase de référence": gold,
                "Sortie du modèle": output,
                "Correspondance exacte normalisée": exact_sentence,
                "Correspondance exacte des éditions":
                    exact_edit_recomputed,
                "Distance de référence": as_number(row["Distance_gold"]),
                "Distance du modèle": as_number(row[config["distance"]]),
                "Nombre d'éditions de référence": gold_count,
                "Nombre d'éditions du modèle": model_count,
                "Éditions de référence recalculées":
                    " | ".join(gold_edit_list),
                "Éditions du modèle recalculées":
                    " | ".join(predicted_edit_list),
                "Éditions de référence — alignement initial":
                    as_text(row["Edits_gold"]),
                "Éditions du modèle — alignement initial":
                    as_text(row[config["edits"]]),
                "Nombre d'éditions initial — référence":
                    as_number(row["Nb_edits_gold"]),
                "Nombre d'éditions initial — modèle":
                    as_number(row[config["n_edits"]]),
                "Correspondance initiale des éditions":
                    as_bool(row[config["exact_edit"]]),
                "Vrais positifs — éditions": metrics["TP_edition"],
                "Faux positifs — éditions": metrics["FP_edition"],
                "Faux négatifs — éditions": metrics["FN_edition"],
                "Précision — éditions": metrics["Precision_edition"],
                "Rappel — éditions": metrics["Rappel_edition"],
                "Score F1 — éditions": metrics["F1_edition"],
                "Diagnostic automatique": automatic_diagnosis(
                    source=source,
                    output=output,
                    exact_sentence=exact_sentence,
                    gold_count=gold_count,
                    model_count=model_count,
                ),
                "Validation manuelle": annotation.get(
                    "Validation manuelle", ""
                ),
                "Commentaire humain": annotation.get(
                    "Commentaire humain", ""
                ),
            })

    return pd.DataFrame(rows)



def micro_edit_metrics(group: pd.DataFrame) -> Dict[str, float]:
    tp = float(group["Vrais positifs — éditions"].sum())
    fp = float(group["Faux positifs — éditions"].sum())
    fn = float(group["Faux négatifs — éditions"].sum())
    precision = safe_ratio(tp, tp + fp)
    recall = safe_ratio(tp, tp + fn)
    f1 = safe_ratio(2 * precision * recall, precision + recall)

    return {
        "TP": tp,
        "FP": fp,
        "FN": fn,
        "Précision": precision,
        "Rappel": recall,
        "F1": f1,
    }


def global_summary(detail: pd.DataFrame) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []

    for model, group in detail.groupby("Modèle", sort=False):
        total = len(group)
        exact = int(group["Correspondance exacte normalisée"].sum())
        exact_edits = int(
            group["Correspondance exacte des éditions"].sum()
        )
        diagnosis = group["Diagnostic automatique"]
        metrics = micro_edit_metrics(group)

        rows.append({
            "Modèle": model,
            "Nombre total de phrases": total,
            "Corrections exactes normalisées": exact,
            "Exactitude normalisée": safe_ratio(exact, total),
            "Correspondances exactes des éditions": exact_edits,
            "Taux de correspondance exacte des éditions":
                safe_ratio(exact_edits, total),
            "Vrais positifs — éditions": metrics["TP"],
            "Faux positifs — éditions": metrics["FP"],
            "Faux négatifs — éditions": metrics["FN"],
            "Précision micro — éditions": metrics["Précision"],
            "Rappel micro — éditions": metrics["Rappel"],
            "Score F1 micro — éditions": metrics["F1"],
            "Précision macro — éditions":
                group["Précision — éditions"].mean(),
            "Rappel macro — éditions":
                group["Rappel — éditions"].mean(),
            "Score F1 macro — éditions":
                group["Score F1 — éditions"].mean(),
            "Absences de modification": int(
                (diagnosis == "Sous-correction : aucune modification").sum()
            ),
            "Sous-corrections probables": int(
                (diagnosis == "Sous-correction probable").sum()
            ),
            "Taux de sous-correction": safe_ratio(
                diagnosis.isin([
                    "Sous-correction : aucune modification",
                    "Sous-correction probable",
                ]).sum(),
                total,
            ),
            "Sur-corrections probables": int(
                (diagnosis == "Sur-correction probable").sum()
            ),
            "Taux de sur-correction": safe_ratio(
                (diagnosis == "Sur-correction probable").sum(),
                total,
            ),
            "Corrections différentes à vérifier": int(
                (diagnosis == "Correction différente à vérifier").sum()
            ),
            "Sorties vides": int((diagnosis == "Sortie vide").sum()),
            "Distance moyenne du modèle":
                group["Distance du modèle"].mean(),
            "Nombre moyen d'éditions du modèle":
                group["Nombre d'éditions du modèle"].mean(),
        })

    return pd.DataFrame(rows)


def category_summary(detail: pd.DataFrame) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []

    for (category, model), group in detail.groupby(
        ["Catégorie", "Modèle"],
        sort=False,
    ):
        total = len(group)
        exact = int(group["Correspondance exacte normalisée"].sum())
        diagnosis = group["Diagnostic automatique"]
        metrics = micro_edit_metrics(group)

        rows.append({
            "Catégorie": category,
            "Modèle": model,
            "Nombre total de phrases": total,
            "Corrections exactes": exact,
            "Exactitude": safe_ratio(exact, total),
            "Précision micro — éditions": metrics["Précision"],
            "Rappel micro — éditions": metrics["Rappel"],
            "Score F1 micro — éditions": metrics["F1"],
            "Sous-corrections": int(
                diagnosis.isin([
                    "Sous-correction : aucune modification",
                    "Sous-correction probable",
                ]).sum()
            ),
            "Sur-corrections": int(
                (diagnosis == "Sur-correction probable").sum()
            ),
            "Corrections différentes": int(
                (diagnosis == "Correction différente à vérifier").sum()
            ),
            "Distance moyenne":
                group["Distance du modèle"].mean(),
            "Nombre moyen d'éditions":
                group["Nombre d'éditions du modèle"].mean(),
        })

    return pd.DataFrame(rows)


def manual_statistics(detail: pd.DataFrame) -> pd.DataFrame:
    reviewed = detail[
        detail["Validation manuelle"].isin(MANUAL_VALUES)
    ].copy()

    columns = [
        "Modèle",
        "Cas validés manuellement",
        "Corrections acceptables",
        "Taux d'acceptabilité humaine",
        "Corrections incorrectes",
        "Sous-corrections confirmées",
        "Sur-corrections confirmées",
        "Corrections hors cible",
        "Cas à discuter",
    ]

    if reviewed.empty:
        return pd.DataFrame(columns=columns)

    rows: List[Dict[str, Any]] = []
    for model, group in reviewed.groupby("Modèle", sort=False):
        counts = group["Validation manuelle"].value_counts()
        total = len(group)
        acceptable = int(counts.get("Correction acceptable", 0))

        rows.append({
            "Modèle": model,
            "Cas validés manuellement": total,
            "Corrections acceptables": acceptable,
            "Taux d'acceptabilité humaine":
                safe_ratio(acceptable, total),
            "Corrections incorrectes":
                int(counts.get("Correction incorrecte", 0)),
            "Sous-corrections confirmées":
                int(counts.get("Sous-correction", 0)),
            "Sur-corrections confirmées":
                int(counts.get("Sur-correction", 0)),
            "Corrections hors cible":
                int(counts.get("Correction hors cible", 0)),
            "Cas à discuter":
                int(counts.get("À discuter", 0)),
        })

    return pd.DataFrame(rows, columns=columns)


def methodology() -> pd.DataFrame:
    return pd.DataFrame([
        {
            "Indicateur": "Exactitude normalisée",
            "Définition": (
                "Proportion de sorties identiques à la correction de référence "
                "après une normalisation légère des espaces, apostrophes, "
                "guillemets et ponctuation finale."
            ),
        },
        {
            "Indicateur": "Précision au niveau des éditions",
            "Définition": (
                "Parmi les éditions produites par le modèle, proportion de "
                "celles qui correspondent à une édition de référence."
            ),
        },
        {
            "Indicateur": "Rappel au niveau des éditions",
            "Définition": (
                "Parmi les éditions attendues dans la référence, proportion de "
                "celles retrouvées dans la sortie du modèle."
            ),
        },
        {
            "Indicateur": "Score F1 au niveau des éditions",
            "Définition": (
                "Moyenne harmonique de la précision et du rappel au niveau "
                "des éditions."
            ),
        },
        {
            "Indicateur": "Micro-moyenne",
            "Définition": (
                "Les vrais positifs, faux positifs et faux négatifs sont "
                "additionnés sur tout le corpus avant le calcul."
            ),
        },
        {
            "Indicateur": "Macro-moyenne",
            "Définition": (
                "La métrique est calculée phrase par phrase, puis moyennée. "
                "Chaque phrase possède ainsi le même poids."
            ),
        },
        {
            "Indicateur": "Validation manuelle",
            "Définition": (
                "Contrôle humain des corrections différentes de la référence. "
                "Une correction peut être acceptable sans être strictement "
                "identique à la phrase gold."
            ),
        },
        {
            "Indicateur": "Limite importante",
            "Définition": (
                "La précision, le rappel et le F1 calculés ici portent sur les "
                "éditions, et non sur la détection binaire de phrases erronées. "
                "Le corpus ne contient pas de véritables négatifs, puisque "
                "toutes les phrases sources sont problématiques."
            ),
        },
        {
            "Indicateur": "Unité d'édition V4",
            "Définition": (
                "Les éditions sont recalculées directement entre la phrase "
                "erronée et la phrase cible. Une différence continue est "
                "comptée comme une suppression, une insertion ou un seul "
                "remplacement, avec sa position dans la phrase source."
            ),
        },
        {
            "Indicateur": "Source des métriques d'édition",
            "Définition": (
                "Les TP, FP, FN, précision, rappel et F1 ne sont plus calculés "
                "à partir des colonnes Edits_* produites par l'alignement "
                "initial. Ces colonnes sont conservées uniquement pour audit."
            ),
        },
        {
            "Indicateur": "Limite du parseur d'éditions",
            "Définition": (
                "La comparaison dépend du format et de l'ordre des colonnes "
                "Edits_*. Le programme prend en charge JSON, listes Python, "
                "lignes et séparateurs courants. Quelques cas doivent rester "
                "contrôlés dans la feuille Evaluation_detaillee."
            ),
        },
    ])


def create_empty_charts_sheet(writer: pd.ExcelWriter) -> None:
    pd.DataFrame({
        "Graphiques": [
            "Les graphiques sont insérés automatiquement par openpyxl."
        ]
    }).to_excel(writer, sheet_name=SHEET_CHARTS, index=False)


def add_manual_validation(workbook) -> None:
    sheet = workbook[SHEET_DETAIL]
    headers = {
        as_text(cell.value): cell.column
        for cell in sheet[1]
    }
    validation_column = headers.get("Validation manuelle")
    if not validation_column:
        return

    formula = '"' + ",".join(MANUAL_VALUES) + '"'
    validation = DataValidation(
        type="list",
        formula1=formula,
        allow_blank=True,
    )
    validation.error = "Choisissez une valeur dans la liste."
    validation.errorTitle = "Valeur non valide"
    validation.prompt = "Sélectionnez le résultat du contrôle humain."
    validation.promptTitle = "Validation manuelle"

    sheet.add_data_validation(validation)
    letter = get_column_letter(validation_column)
    validation.add(f"{letter}2:{letter}{max(sheet.max_row, 2)}")



def add_chart(
    chart_sheet,
    data_sheet,
    title: str,
    categories_column: int,
    data_columns: Sequence[int],
    anchor: str,
    y_title: str,
    percent_axis: bool = False,
) -> bool:
    """Ajoute un graphique uniquement si les données nécessaires existent."""
    if data_sheet.max_row < 2:
        return False

    valid_columns = [
        column for column in data_columns
        if isinstance(column, int) and 1 <= column <= data_sheet.max_column
    ]
    if not valid_columns:
        return False

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = title
    chart.y_axis.title = y_title
    chart.x_axis.title = "Modèle"
    chart.height = 8
    chart.width = 15
    chart.grouping = "clustered"
    chart.overlap = 0
    chart.legend.position = "r"

    categories = Reference(
        data_sheet,
        min_col=categories_column,
        min_row=2,
        max_row=data_sheet.max_row,
    )

    for column in valid_columns:
        data = Reference(
            data_sheet,
            min_col=column,
            min_row=1,
            max_row=data_sheet.max_row,
        )
        chart.add_data(data, titles_from_data=True)

    chart.set_categories(categories)
    chart.dataLabels = DataLabelList()
    chart.dataLabels.showVal = True

    if percent_axis:
        chart.y_axis.numFmt = "0%"
        chart.y_axis.scaling.min = 0
        chart.y_axis.scaling.max = 1

    chart_sheet.add_chart(chart, anchor)
    return True


def write_chart_source_table(chart_sheet, global_sheet) -> None:
    """
    Copie un tableau de synthèse sous les graphiques.

    Cette table permet de contrôler les valeurs utilisées par les graphiques
    et reste visible même dans un lecteur Excel qui ne rend pas les objets
    graphiques openpyxl.
    """
    headers = {
        as_text(cell.value): cell.column
        for cell in global_sheet[1]
    }
    selected = [
        "Modèle",
        "Exactitude normalisée",
        "Précision micro — éditions",
        "Rappel micro — éditions",
        "Score F1 micro — éditions",
        "Absences de modification",
        "Sous-corrections probables",
        "Sur-corrections probables",
    ]
    selected = [header for header in selected if header in headers]

    start_row = 39
    chart_sheet.cell(start_row, 1, "Données utilisées pour les graphiques")
    chart_sheet.cell(start_row, 1).font = Font(bold=True, size=12)

    for column_index, header in enumerate(selected, start=1):
        cell = chart_sheet.cell(start_row + 1, column_index, header)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    for row_index in range(2, global_sheet.max_row + 1):
        for column_index, header in enumerate(selected, start=1):
            value = global_sheet.cell(
                row=row_index,
                column=headers[header],
            ).value
            target = chart_sheet.cell(
                row=start_row + row_index,
                column=column_index,
                value=value,
            )
            if any(token in header for token in (
                "Exactitude", "Précision", "Rappel", "Score F1"
            )):
                target.number_format = "0.00%"

    for column_index, header in enumerate(selected, start=1):
        chart_sheet.column_dimensions[
            get_column_letter(column_index)
        ].width = max(16, min(len(header) + 3, 34))


def add_charts(workbook) -> None:
    """Crée les graphiques du mémoire à partir des feuilles de résultats."""
    chart_sheet = workbook[SHEET_CHARTS]

    # Supprime les anciens graphiques et le contenu textuel.
    chart_sheet._charts = []
    if chart_sheet.max_row:
        chart_sheet.delete_rows(1, chart_sheet.max_row)

    chart_sheet.sheet_view.showGridLines = False
    chart_sheet["A1"] = "Graphiques pour le mémoire"
    chart_sheet["A1"].font = Font(size=16, bold=True)
    chart_sheet["A2"] = (
        "Les graphiques sont régénérés à chaque exécution. "
        "Le tableau de contrôle situé sous les graphiques reprend leurs données."
    )
    chart_sheet["A2"].alignment = Alignment(wrap_text=True)
    chart_sheet.column_dimensions["A"].width = 50

    global_sheet = workbook[SHEET_GLOBAL]
    global_headers = {
        as_text(cell.value): cell.column
        for cell in global_sheet[1]
    }

    required_global = {
        "Modèle",
        "Exactitude normalisée",
        "Précision micro — éditions",
        "Rappel micro — éditions",
        "Score F1 micro — éditions",
        "Absences de modification",
        "Sous-corrections probables",
        "Sur-corrections probables",
    }
    missing = sorted(required_global - set(global_headers))
    if missing:
        chart_sheet["A4"] = (
            "Graphiques globaux non générés. Colonnes manquantes : "
            + ", ".join(missing)
        )
    else:
        add_chart(
            chart_sheet=chart_sheet,
            data_sheet=global_sheet,
            title="Exactitude normalisée par modèle",
            categories_column=global_headers["Modèle"],
            data_columns=[global_headers["Exactitude normalisée"]],
            anchor="A4",
            y_title="Exactitude",
            percent_axis=True,
        )

        add_chart(
            chart_sheet=chart_sheet,
            data_sheet=global_sheet,
            title="Précision, rappel et score F1 au niveau des éditions",
            categories_column=global_headers["Modèle"],
            data_columns=[
                global_headers["Précision micro — éditions"],
                global_headers["Rappel micro — éditions"],
                global_headers["Score F1 micro — éditions"],
            ],
            anchor="Q4",
            y_title="Score",
            percent_axis=True,
        )

        add_chart(
            chart_sheet=chart_sheet,
            data_sheet=global_sheet,
            title="Sous-corrections et sur-corrections probables",
            categories_column=global_headers["Modèle"],
            data_columns=[
                global_headers["Absences de modification"],
                global_headers["Sous-corrections probables"],
                global_headers["Sur-corrections probables"],
            ],
            anchor="A21",
            y_title="Nombre de phrases",
            percent_axis=False,
        )

        write_chart_source_table(chart_sheet, global_sheet)

    manual_sheet = workbook[SHEET_MANUAL]
    if manual_sheet.max_row >= 2:
        manual_headers = {
            as_text(cell.value): cell.column
            for cell in manual_sheet[1]
        }
        if (
            "Modèle" in manual_headers
            and "Taux d'acceptabilité humaine" in manual_headers
        ):
            add_chart(
                chart_sheet=chart_sheet,
                data_sheet=manual_sheet,
                title="Taux d'acceptabilité après validation humaine",
                categories_column=manual_headers["Modèle"],
                data_columns=[
                    manual_headers["Taux d'acceptabilité humaine"]
                ],
                anchor="Q21",
                y_title="Taux d'acceptabilité",
                percent_axis=True,
            )




def style_workbook(path: Path) -> None:
    """Met en forme le classeur, puis ajoute validations et graphiques."""
    workbook = load_workbook(path)

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    thin_gray = Side(style="thin", color="D9D9D9")
    border = Border(
        left=thin_gray,
        right=thin_gray,
        top=thin_gray,
        bottom=thin_gray,
    )

    for sheet in workbook.worksheets:
        if sheet.title == SHEET_CHARTS:
            continue

        sheet.freeze_panes = "A2"
        sheet.sheet_view.showGridLines = False

        # Évite d'appliquer un filtre à une feuille vide.
        if sheet.max_row >= 1 and sheet.max_column >= 1:
            sheet.auto_filter.ref = sheet.dimensions

        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = Font(bold=True)
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True,
            )
            cell.border = border

        for row in sheet.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(
                    vertical="top",
                    wrap_text=True,
                )
                cell.border = border

        headers = {
            as_text(cell.value): cell.column
            for cell in sheet[1]
            if as_text(cell.value)
        }

        for header, column in headers.items():
            letter = get_column_letter(column)

            if any(token in header for token in (
                "Phrase",
                "Sortie",
                "Éditions",
                "Corrections",
                "Commentaire",
                "Définition",
            )):
                width = 42
            elif header in {
                "Diagnostic automatique",
                "Validation manuelle",
            }:
                width = 29
            else:
                values = [
                    len(as_text(sheet.cell(row=row, column=column).value))
                    for row in range(1, min(sheet.max_row, 100) + 1)
                ]
                width = min(max(max(values, default=10) + 2, 12), 28)

            sheet.column_dimensions[letter].width = width

            if any(token in header for token in (
                "Exactitude",
                "Précision",
                "Rappel",
                "Score F1",
                "Taux",
            )):
                for row_index in range(2, sheet.max_row + 1):
                    cell = sheet.cell(row=row_index, column=column)
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "0.00%"

        if (
            sheet.title in {
                SHEET_GLOBAL,
                SHEET_CATEGORY,
                SHEET_MATRIX,
                SHEET_MANUAL,
            }
            and sheet.max_row >= 2
        ):
            percentage_columns = [
                column
                for header, column in headers.items()
                if any(token in header for token in (
                    "Exactitude",
                    "Précision",
                    "Rappel",
                    "Score F1",
                    "Taux",
                ))
            ]

            for column in percentage_columns:
                letter = get_column_letter(column)
                valid_range = f"{letter}2:{letter}{sheet.max_row}"
                sheet.conditional_formatting.add(
                    valid_range,
                    ColorScaleRule(
                        start_type="num",
                        start_value=0,
                        start_color="F8696B",
                        mid_type="num",
                        mid_value=0.5,
                        mid_color="FFEB84",
                        end_type="num",
                        end_value=1,
                        end_color="63BE7B",
                    ),
                )

    add_manual_validation(workbook)
    add_charts(workbook)
    workbook.save(path)



def main() -> None:
    args = parse_args()
    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        raise FileNotFoundError(
            f"Fichier d'entrée introuvable : {input_path.resolve()}"
        )

    output_path.parent.mkdir(parents=True, exist_ok=True)

    annotations: Dict[Tuple[str, str], Dict[str, str]] = {}
    if not args.no_preserve_manual:
        annotations = load_manual_annotations(output_path)

    source_df = pd.read_excel(input_path)
    validate_columns(source_df)

    detail = build_detail(source_df, annotations)
    global_df = global_summary(detail)
    category_df = category_summary(detail)
    manual_df = manual_statistics(detail)

    matrix_df = category_df.pivot(
        index="Catégorie",
        columns="Modèle",
        values="Exactitude",
    ).reset_index()

    review_df = detail[
        ~detail["Correspondance exacte normalisée"]
    ].copy()

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        global_df.to_excel(
            writer,
            sheet_name=SHEET_GLOBAL,
            index=False,
        )
        category_df.to_excel(
            writer,
            sheet_name=SHEET_CATEGORY,
            index=False,
        )
        matrix_df.to_excel(
            writer,
            sheet_name=SHEET_MATRIX,
            index=False,
        )
        detail.to_excel(
            writer,
            sheet_name=SHEET_DETAIL,
            index=False,
        )
        review_df.to_excel(
            writer,
            sheet_name=SHEET_REVIEW,
            index=False,
        )
        manual_df.to_excel(
            writer,
            sheet_name=SHEET_MANUAL,
            index=False,
        )
        create_empty_charts_sheet(writer)
        source_df.to_excel(
            writer,
            sheet_name=SHEET_SOURCE,
            index=False,
        )
        methodology().to_excel(
            writer,
            sheet_name=SHEET_METHOD,
            index=False,
        )

    style_workbook(output_path)

    reviewed_count = int(
        detail["Validation manuelle"].isin(MANUAL_VALUES).sum()
    )

    print(f"Évaluation terminée : {output_path}")
    print(f"Nombre de phrases : {len(source_df)}")
    print(f"Nombre de modèles : {len(MODELS)}")
    print(f"Validations manuelles conservées : {reviewed_count}")


if __name__ == "__main__":
    main()
