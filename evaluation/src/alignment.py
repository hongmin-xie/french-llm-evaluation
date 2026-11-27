from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

import pandas as pd


SOURCE_COLUMN = "Phrase_erronée"

TARGET_COLUMNS = [
    "Phrase_gold",
    "Sortie_gpt_oss_20b",
    "Sortie_deepseek_v4_flash",
    "Sortie_qwen3_32b",
    "Sortie_gemma3_27b",
]


def load_stanza_json(path: Path) -> list[dict[str, Any]]:
    return json.loads(path.read_text(encoding="utf-8"))


def token_cost(orig_token: dict[str, Any], corr_token: dict[str, Any]) -> int:
    """Attribue un coût nul aux tokens identiques et un coût unitaire aux substitutions.
    
    L’alignement repose sur la programmation dynamique, avec correspondances,
    substitutions, insertions et suppressions."""
    return 0 if orig_token["text"] == corr_token["text"] else 1


def align_tokens(
    orig_tokens: list[dict[str, Any]],
    corr_tokens: list[dict[str, Any]],
) -> dict[str, Any]:
    """Calcule un alignement de Levenshtein au niveau des tokens par programmation dynamique.
    
    Renvoie :
    - operations : toutes les opérations, correspondances comprises ;
    - edits : substitutions, insertions et suppressions uniquement ;
    - distance : distance d’édition minimale."""
    n = len(orig_tokens)
    m = len(corr_tokens)

    dp = [[0] * (m + 1) for _ in range(n + 1)]
    back = [[None] * (m + 1) for _ in range(n + 1)]

    for i in range(1, n + 1):
        dp[i][0] = i
        back[i][0] = "D"

    for j in range(1, m + 1):
        dp[0][j] = j
        back[0][j] = "I"

    for i in range(1, n + 1):
        for j in range(1, m + 1):
            substitution_cost = token_cost(orig_tokens[i - 1], corr_tokens[j - 1])

            candidates = [
                (
                    dp[i - 1][j - 1] + substitution_cost,
                    "M" if substitution_cost == 0 else "R",
                ),
                (dp[i - 1][j] + 1, "D"),
                (dp[i][j - 1] + 1, "I"),
            ]

            # Ordre de priorité en cas de coûts identiques :
            # 1. Match
            # 2. Replacement
            # 3. Deletion
            # 4. Insertion
            priority = {"M": 0, "R": 1, "D": 2, "I": 3}
            best_cost, best_op = min(
                candidates,
                key=lambda item: (item[0], priority[item[1]]),
            )

            dp[i][j] = best_cost
            back[i][j] = best_op

    operations: list[dict[str, Any]] = []
    i, j = n, m

    while i > 0 or j > 0:
        op = back[i][j]

        if op in {"M", "R"}:
            orig = orig_tokens[i - 1]
            corr = corr_tokens[j - 1]

            operations.append(
                {
                    "type": "Match" if op == "M" else "Replacement",
                    "orig_index": i - 1,
                    "corr_index": j - 1,
                    "orig_token": orig,
                    "corr_token": corr,
                }
            )
            i -= 1
            j -= 1

        elif op == "D":
            orig = orig_tokens[i - 1]

            operations.append(
                {
                    "type": "Deletion",
                    "orig_index": i - 1,
                    "corr_index": None,
                    "orig_token": orig,
                    "corr_token": None,
                }
            )
            i -= 1

        elif op == "I":
            corr = corr_tokens[j - 1]

            operations.append(
                {
                    "type": "Insertion",
                    "orig_index": None,
                    "corr_index": j - 1,
                    "orig_token": None,
                    "corr_token": corr,
                }
            )
            j -= 1

        else:
            raise RuntimeError(
                f"无法回溯对齐路径：i={i}, j={j}, op={op}"
            )

    operations.reverse()

    edits = [
        operation
        for operation in operations
        if operation["type"] != "Match"
    ]

    return {
        "distance": dp[n][m],
        "operations": operations,
        "edits": edits,
    }


def simplify_edits(edits: list[dict[str, Any]]) -> str:
    """Convertit les opérations d’édition en texte concis pour l’export Excel."""
    descriptions: list[str] = []

    for edit in edits:
        edit_type = edit["type"]

        if edit_type == "Replacement":
            orig = edit["orig_token"]["text"]
            corr = edit["corr_token"]["text"]
            descriptions.append(f"Remplacement: {orig} → {corr}")

        elif edit_type == "Deletion":
            orig = edit["orig_token"]["text"]
            descriptions.append(f"Suppression: {orig}")

        elif edit_type == "Insertion":
            corr = edit["corr_token"]["text"]
            descriptions.append(f"Insertion: {corr}")

    return " | ".join(descriptions) if descriptions else "Aucune modification"


def edit_signature(edits: list[dict[str, Any]]) -> list[tuple[str, str | None, str | None]]:
    """Convertit les opérations d’édition en signatures comparables.
    
    Permet de vérifier la correspondance exacte entre les éditions du modèle
    et celles de la référence (gold)."""
    signature = []

    for edit in edits:
        orig = (
            None
            if edit["orig_token"] is None
            else edit["orig_token"]["text"]
        )
        corr = (
            None
            if edit["corr_token"] is None
            else edit["corr_token"]["text"]
        )
        signature.append((edit["type"], orig, corr))

    return signature


def process_alignments(
    stanza_json: Path,
    original_excel: Path,
    output_json: Path,
    output_excel: Path,
) -> None:
    analyses = load_stanza_json(stanza_json)
    df = pd.read_excel(original_excel)

    output_rows: list[dict[str, Any]] = []
    output_docs: list[dict[str, Any]] = []

    for position, item in enumerate(analyses):
        row_id = item.get("id")
        row_analyses = item["analyses"]

        source_tokens = row_analyses[SOURCE_COLUMN]["tokens"]
        source_text = row_analyses[SOURCE_COLUMN]["text"]

        doc_result: dict[str, Any] = {
            "row_index": item["row_index"],
            "id": row_id,
            "source_column": SOURCE_COLUMN,
            "source_text": source_text,
            "alignments": {},
        }

        excel_row: dict[str, Any] = {
            "ID": row_id,
            SOURCE_COLUMN: source_text,
        }

        gold_alignment = align_tokens(
            source_tokens,
            row_analyses["Phrase_gold"]["tokens"],
        )
        gold_signature = edit_signature(gold_alignment["edits"])

        for target_column in TARGET_COLUMNS:
            target_text = row_analyses[target_column]["text"]
            target_tokens = row_analyses[target_column]["tokens"]

            alignment = align_tokens(source_tokens, target_tokens)
            signature = edit_signature(alignment["edits"])

            exact_edit_match = (
                target_column == "Phrase_gold"
                or signature == gold_signature
            )

            doc_result["alignments"][target_column] = {
                "target_text": target_text,
                **alignment,
                "edit_signature": signature,
                "exact_edit_match_gold": exact_edit_match,
            }

            prefix = target_column.replace("Sortie_", "").replace("Phrase_", "")

            excel_row[target_column] = target_text
            excel_row[f"Distance_{prefix}"] = alignment["distance"]
            excel_row[f"Edits_{prefix}"] = simplify_edits(alignment["edits"])
            excel_row[f"Nb_edits_{prefix}"] = len(alignment["edits"])

            if target_column != "Phrase_gold":
                excel_row[f"Exact_edit_match_gold_{prefix}"] = exact_edit_match

        output_docs.append(doc_result)
        output_rows.append(excel_row)

        print(f"[{position + 1}/{len(analyses)}] Alignement, ID={row_id}")

        output_json.parent.mkdir(parents=True, exist_ok=True)
        output_json.write_text(
            json.dumps(output_docs, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    result_df = pd.DataFrame(output_rows)

    # Conserver les colonnes de métadonnées du classeur source.
    metadata_columns = [
        col
        for col in ["ID", "Source", "Habileté", "Corrections"]
        if col in df.columns
    ]

    metadata = df[metadata_columns].copy()

    if "ID" in metadata.columns and "ID" in result_df.columns:
        final_df = metadata.merge(result_df, on="ID", how="right")
    else:
        final_df = pd.concat(
            [metadata.reset_index(drop=True), result_df.reset_index(drop=True)],
            axis=1,
        )

    output_excel.parent.mkdir(parents=True, exist_ok=True)
    final_df.to_excel(output_excel, index=False)

    print(f"\n对齐 JSON：{output_json}")
    print(f"对齐 Excel：{output_excel}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="根据 Stanza token 分析结果，对齐原句、Gold 和四个 LLM 输出。"
    )
    parser.add_argument(
        "--stanza-json",
        type=Path,
        required=True,
        help="stanza_analysis.py 生成的 JSON",
    )
    parser.add_argument(
        "--input-excel",
        type=Path,
        required=True,
        help="原始 corpus complet.xlsx",
    )
    parser.add_argument(
        "--output-json",
        type=Path,
        default=Path("result/alignments.json"),
        help="完整对齐 JSON",
    )
    parser.add_argument(
        "--output-excel",
        type=Path,
        default=Path("result/alignments.xlsx"),
        help="便于人工查看的对齐 Excel",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()

    process_alignments(
        stanza_json=args.stanza_json,
        original_excel=args.input_excel,
        output_json=args.output_json,
        output_excel=args.output_excel,
    )
