import os
import time
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


# ============================================================
# 1. Chemins du projet
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parent.parent

INPUT_FILE = PROJECT_ROOT / "data" / "corpus_107.xlsx"

OUTPUT_FILE = (
    PROJECT_ROOT
    / "result"
    / "results_107_gpt_oss_20b.xlsx"
)

PROMPT_FILE = (
    PROJECT_ROOT
    / "prompts"
    / "prompt_v1.txt"
)


# ============================================================
# 2. Configuration du modèle et paramètres expérimentaux
# ============================================================

MODEL_NAME = "openai/gpt-oss-20b"

INPUT_COLUMN_NAME = "Phrase erronée"

TEMPERATURE = 0
MAX_TOKENS = 1000

# Temporisation de deux secondes entre deux requêtes réussies.
REQUEST_DELAY = 2

# Nombre de tentatives supplémentaires en cas de réponse vide ou d’erreur API.
# La valeur 1 autorise une nouvelle tentative après le premier échec.
MAX_RETRIES = 1

# En cas de quota épuisé (402), réessayer la phrase courante toutes les 20 secondes.
CREDIT_CHECK_INTERVAL = 20


# ============================================================
# 3. Fonctions utilitaires pour Excel
# ============================================================

def find_column(
    worksheet: Worksheet,
    column_name: str,
) -> int:
    """Recherche un nom de colonne dans la première ligne du classeur Excel
    et renvoie son indice."""

    for cell in worksheet[1]:
        if cell.value is None:
            continue

        current_name = str(cell.value).strip()

        if current_name == column_name:
            return cell.column

    available_columns = [
        str(cell.value).strip()
        for cell in worksheet[1]
        if cell.value is not None
    ]

    raise ValueError(
        f'找不到列 "{column_name}"。\n'
        f"当前Excel中的列名为：{available_columns}"
    )


def get_or_create_column(
    worksheet: Worksheet,
    column_name: str,
) -> int:
    """Renvoie l’indice de la colonne existante ou crée une colonne
    à la fin du classeur."""

    for cell in worksheet[1]:
        if cell.value is None:
            continue

        current_name = str(cell.value).strip()

        if current_name == column_name:
            return cell.column

    new_column = worksheet.max_column + 1

    worksheet.cell(
        row=1,
        column=new_column,
        value=column_name,
    )

    return new_column


# ============================================================
# 4. Identification des erreurs
# ============================================================

def is_credit_error(error_message: str) -> bool:
    """Identifie une erreur d’épuisement du quota Hugging Face."""

    message = str(error_message).lower()

    return (
        "error code: 402" in message
        or "depleted your monthly included credits" in message
        or "purchase pre-paid credits" in message
    )


# ============================================================
# 5. Appel au modèle
# ============================================================

def call_model(
    client: OpenAI,
    system_prompt: str,
    sentence: str,
) -> dict[str, Any]:
    """Interroge le modèle via Hugging Face.
    
    Une réponse non vide est renvoyée directement. Les réponses vides et les
    erreurs API ordinaires autorisent au plus MAX_RETRIES tentatives supplémentaires.
    En cas d’erreur 402, la phrase courante est retraitée toutes les 20 secondes
    jusqu’à réussite ou interruption par Ctrl+C."""

    normal_attempts = 0
    credit_waits = 0
    total_request_time = 0.0

    while True:
        start_time = time.perf_counter()

        try:
            response = client.chat.completions.create(
                model=MODEL_NAME,
                messages=[
                    {
                        "role": "system",
                        "content": system_prompt,
                    },
                    {
                        "role": "user",
                        "content": (
                            f"Phrase à corriger : {sentence}"
                        ),
                    },
                ],
                temperature=TEMPERATURE,
                max_tokens=MAX_TOKENS,
            )

            elapsed_time = (
                time.perf_counter() - start_time
            )

            total_request_time += elapsed_time

            output = (
                response
                .choices[0]
                .message
                .content
            )

            if output is None:
                output = ""
            else:
                output = output.strip()

            if output:
                return {
                    "output": output,
                    "status": "OK",
                    "time": total_request_time,
                    "attempts": normal_attempts + 1,
                    "credit_waits": credit_waits,
                }

            # Réponse vide du modèle.
            normal_attempts += 1

            print(
                f"第 {normal_attempts} 次调用返回空输出。"
            )

            if normal_attempts > MAX_RETRIES:
                return {
                    "output": "",
                    "status": "FAILED: 模型返回空输出",
                    "time": total_request_time,
                    "attempts": normal_attempts,
                    "credit_waits": credit_waits,
                }

            print("等待3秒后重试当前句……")
            time.sleep(3)

        except KeyboardInterrupt:
            print("\n收到 Ctrl+C，程序即将停止。")
            raise

        except Exception as error:
            elapsed_time = (
                time.perf_counter() - start_time
            )

            total_request_time += elapsed_time

            error_text = (
                f"{type(error).__name__}: {error}"
            )

            # ------------------------------------------------
            # 402 : quota épuisé
            # ------------------------------------------------

            if is_credit_error(error_text):
                credit_waits += 1

                print("\n" + "!" * 70)
                print("检测到 Hugging Face 额度不足。")
                print("20秒后重新请求当前句。")
                print(
                    f"当前额度等待次数：{credit_waits}"
                )
                print(
                    "按 Ctrl+C 可以停止程序。"
                )
                print("!" * 70)

                time.sleep(CREDIT_CHECK_INTERVAL)

                # Réessayer la phrase courante sans passer à la suivante.
                continue

            # ------------------------------------------------
            # Autres erreurs API
            # ------------------------------------------------

            normal_attempts += 1

            print(
                f"第 {normal_attempts} 次调用失败："
            )
            print(error_text)

            if normal_attempts > MAX_RETRIES:
                return {
                    "output": "",
                    "status": f"FAILED: {error_text}",
                    "time": total_request_time,
                    "attempts": normal_attempts,
                    "credit_waits": credit_waits,
                }

            print("等待3秒后重试当前句……")
            time.sleep(3)


# ============================================================
# 6. Point d’entrée
# ============================================================

def main() -> None:
    """Traite par lots les phrases du classeur Excel."""

    # --------------------------------------------------------
    # Chargement des variables d’environnement depuis .env.
    # --------------------------------------------------------

    load_dotenv(PROJECT_ROOT / ".env")

    hf_token = os.getenv("HF_TOKEN")

    if not hf_token:
        raise RuntimeError(
            "没有读取到 HF_TOKEN。\n"
            "请检查项目根目录下的 .env 文件。"
        )

    # --------------------------------------------------------
    # Vérification des fichiers.
    # --------------------------------------------------------

    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"找不到输入文件：\n{INPUT_FILE}\n\n"
            "请确认完整语料文件名为："
            "corpus_107.xlsx"
        )

    if not PROMPT_FILE.exists():
        raise FileNotFoundError(
            f"找不到Prompt文件：\n{PROMPT_FILE}"
        )

    OUTPUT_FILE.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    # --------------------------------------------------------
    # Lecture du prompt.
    # --------------------------------------------------------

    system_prompt = PROMPT_FILE.read_text(
        encoding="utf-8"
    ).strip()

    if not system_prompt:
        raise ValueError(
            f"Prompt文件为空：\n{PROMPT_FILE}"
        )

    # --------------------------------------------------------
    # Initialisation du client Hugging Face.
    # --------------------------------------------------------

    client = OpenAI(
        base_url="https://router.huggingface.co/v1",
        api_key=hf_token,
    )

    # --------------------------------------------------------
    # Lecture du classeur de résultats ou du corpus source.
    # --------------------------------------------------------

    if OUTPUT_FILE.exists():
        print("检测到已有结果文件。")
        print("程序将继续上一次实验。")
        print("读取文件：", OUTPUT_FILE)

        workbook = load_workbook(OUTPUT_FILE)

    else:
        print("没有检测到旧结果文件。")
        print("程序将开始新的实验。")
        print("读取文件：", INPUT_FILE)

        workbook = load_workbook(INPUT_FILE)

    worksheet = workbook.active

    # --------------------------------------------------------
    # Identification de la colonne d’entrée.
    # --------------------------------------------------------

    input_column = find_column(
        worksheet,
        INPUT_COLUMN_NAME,
    )

    # --------------------------------------------------------
    # Création des colonnes de résultats.
    # --------------------------------------------------------

    output_column = get_or_create_column(
        worksheet,
        "Sortie_gpt_oss_20b",
    )

    status_column = get_or_create_column(
        worksheet,
        "Statut_gpt_oss_20b",
    )

    time_column = get_or_create_column(
        worksheet,
        "Temps_secondes_gpt_oss_20b",
    )

    attempts_column = get_or_create_column(
        worksheet,
        "Tentatives_gpt_oss_20b",
    )

    credit_waits_column = get_or_create_column(
        worksheet,
        "Attentes_crédit_gpt_oss_20b",
    )

    model_column = get_or_create_column(
        worksheet,
        "Modèle_testé",
    )

    # --------------------------------------------------------
    # Initialisation des compteurs.
    # --------------------------------------------------------

    total_rows = worksheet.max_row - 1

    processed_count = 0
    skipped_count = 0
    failed_count = 0

    print("\n" + "=" * 70)
    print("模型：", MODEL_NAME)
    print("语料行数：", total_rows)
    print("结果文件：", OUTPUT_FILE)
    print("=" * 70)

    # --------------------------------------------------------
    # Traitement à partir de la deuxième ligne du classeur.
    # --------------------------------------------------------

    try:
        for row_number in range(
            2,
            worksheet.max_row + 1,
        ):
            sentence_value = worksheet.cell(
                row=row_number,
                column=input_column,
            ).value

            # Ignorer les lignes vides.
            if sentence_value is None:
                skipped_count += 1
                continue

            sentence = str(sentence_value).strip()

            if not sentence:
                skipped_count += 1
                continue

            # Vérifier si la ligne dispose déjà d’un résultat valide.
            existing_output = worksheet.cell(
                row=row_number,
                column=output_column,
            ).value

            existing_status = worksheet.cell(
                row=row_number,
                column=status_column,
            ).value

            # Ignorer uniquement les lignes dont la sortie est renseignée et le statut est OK.
            # Retraiter les échecs antérieurs, notamment les erreurs 402 et les réponses vides.
            if (
                existing_output
                and existing_status == "OK"
            ):
                print(
                    f"第 {row_number} 行已有成功结果，跳过。"
                )

                skipped_count += 1
                continue

            current_sentence_number = row_number - 1

            print("\n" + "=" * 70)
            print(
                f"正在处理第 "
                f"{current_sentence_number} / "
                f"{total_rows} 句"
            )
            print("Excel行号：", row_number)
            print("输入：", sentence)

            result = call_model(
                client=client,
                system_prompt=system_prompt,
                sentence=sentence,
            )

            output = result["output"]
            status = result["status"]
            elapsed_time = result["time"]
            attempts = result["attempts"]
            credit_waits = result["credit_waits"]

            # ------------------------------------------------
            # Écriture des résultats dans Excel.
            # ------------------------------------------------

            worksheet.cell(
                row=row_number,
                column=output_column,
                value=output,
            )

            worksheet.cell(
                row=row_number,
                column=status_column,
                value=status,
            )

            worksheet.cell(
                row=row_number,
                column=time_column,
                value=round(elapsed_time, 3),
            )

            worksheet.cell(
                row=row_number,
                column=attempts_column,
                value=attempts,
            )

            worksheet.cell(
                row=row_number,
                column=credit_waits_column,
                value=credit_waits,
            )

            worksheet.cell(
                row=row_number,
                column=model_column,
                value=MODEL_NAME,
            )

            processed_count += 1

            if status != "OK":
                failed_count += 1

            # ------------------------------------------------
            # Affichage du résultat courant.
            # ------------------------------------------------

            print("输出：", output)
            print("状态：", status)
            print("普通尝试次数：", attempts)
            print("额度等待次数：", credit_waits)
            print(f"API累计耗时：{elapsed_time:.2f} 秒")

            # Enregistrer le classeur après chaque phrase.
            workbook.save(OUTPUT_FILE)

            print("已保存：", OUTPUT_FILE)

            # Temporisation entre les requêtes.
            time.sleep(REQUEST_DELAY)

    except KeyboardInterrupt:
        print("\n" + "=" * 70)
        print("用户通过 Ctrl+C 停止程序。")
        print("正在保存当前结果……")

        workbook.save(OUTPUT_FILE)

        print("结果已保存：", OUTPUT_FILE)
        print("下次重新运行时会继续未完成的句子。")
        print("=" * 70)

        return

    # --------------------------------------------------------
    # Bilan du traitement.
    # --------------------------------------------------------

    workbook.save(OUTPUT_FILE)

    print("\n" + "=" * 70)
    print("实验运行结束。")
    print("本次新处理：", processed_count)
    print("跳过已有成功结果：", skipped_count)
    print("失败或空输出：", failed_count)
    print("结果文件：", OUTPUT_FILE)
    print("=" * 70)


if __name__ == "__main__":
    main()