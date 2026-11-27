from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

import pandas as pd
import stanza


DEFAULT_COLUMNS = [
    "Phrase_erronée",
    "Phrase_gold",
    "Sortie_gpt_oss_20b",
    "Sortie_deepseek_v4_flash",
    "Sortie_qwen3_32b",
    "Sortie_gemma3_27b",
]


def build_pipeline(lang: str = "fr", use_gpu: bool = False) -> stanza.Pipeline:
    """Initialise le pipeline d’analyse du français avec Stanza.
    
    Champs extraits : text, lemma, upos, xpos, feats, head et deprel."""
    try:
        return stanza.Pipeline(
            lang=lang,
            processors="tokenize,pos,lemma,depparse",
            use_gpu=use_gpu,
            tokenize_no_ssplit=True,
            verbose=False,
        )
    except Exception as exc:
        raise RuntimeError(
            "无法加载 Stanza 法语模型。请先运行：\n"
            "python -c \"import stanza; stanza.download('fr')\""
        ) from exc


def analyse_sentence(nlp: stanza.Pipeline, sentence: str) -> list[dict[str, Any]]:
    """Effectue l’analyse morphosyntaxique d’une phrase avec Stanza."""
    if not isinstance(sentence, str) or not sentence.strip():
        return []

    doc = nlp(sentence.strip())
    tokens: list[dict[str, Any]] = []

    for sent in doc.sentences:
        for word in sent.words:
            tokens.append(
                {
                    "id": int(word.id),
                    "text": word.text,
                    "lemma": word.lemma,
                    "upos": word.upos,
                    "xpos": word.xpos,
                    "feats": word.feats,
                    "head": int(word.head),
                    "deprel": word.deprel,
                }
            )

    return tokens


def process_excel(
    input_path: Path,
    output_json: Path,
    columns: list[str],
    use_gpu: bool = False,
) -> None:
    """Lit le classeur Excel, analyse les phrases des colonnes sélectionnées
    et enregistre les résultats au format JSON."""
    df = pd.read_excel(input_path)

    missing = [col for col in columns if col not in df.columns]
    if missing:
        raise ValueError(
            "Excel 中缺少以下列：\n" + "\n".join(missing)
        )

    nlp = build_pipeline(use_gpu=use_gpu)
    results: list[dict[str, Any]] = []

    total = len(df)

    for index, row in df.iterrows():
        row_id = row.get("ID", index + 1)

        item: dict[str, Any] = {
            "row_index": int(index),
            "id": None if pd.isna(row_id) else row_id,
            "analyses": {},
        }

        print(f"[{index + 1}/{total}] Analyse Stanza, ID={item['id']}")

        for column in columns:
            value = row.get(column)

            if pd.isna(value):
                sentence = ""
            else:
                sentence = str(value).strip()

            item["analyses"][column] = {
                "text": sentence,
                "tokens": analyse_sentence(nlp, sentence),
            }

        results.append(item)

        # Enregistrer après chaque phrase afin de permettre la conservation des résultats en cas d’interruption.
        output_json.parent.mkdir(parents=True, exist_ok=True)
        output_json.write_text(
            json.dumps(results, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    print(f"\nStanza 分析完成：{output_json}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="对论文语料中的原句、Gold 和四个模型输出做 Stanza 分析。"
    )
    parser.add_argument(
        "--input",
        type=Path,
        required=True,
        help="输入 Excel 文件路径",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("result/stanza_analyses.json"),
        help="输出 JSON 文件路径",
    )
    parser.add_argument(
        "--gpu",
        action="store_true",
        help="使用 GPU 运行 Stanza",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()

    process_excel(
        input_path=args.input,
        output_json=args.output,
        columns=DEFAULT_COLUMNS,
        use_gpu=args.gpu,
    )
