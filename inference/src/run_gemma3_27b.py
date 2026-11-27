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

INPUT_FILE = (
    PROJECT_ROOT
    / "data"
    / "corpus_107.xlsx"
)

PROMPT_FILE = (
    PROJECT_ROOT
    / "prompts"
    / "prompt_v1.txt"
)

OUTPUT_FILE = (
    PROJECT_ROOT
    / "result"
    / "results_107_gemma3_27b.xlsx"
)


# ============================================================
# 2. Configuration du modèle
# ============================================================

MODEL_NAME = "google/gemma-3-27b-it"

MODEL_SHORT_NAME = "gemma3_27b"

# Colonne du classeur contenant les phrases erronées.
INPUT_COLUMN_NAME = "Phrase erronée"

# Pour garantir la comparabilité des expériences,
# conserver les paramètres utilisés pour GPT-OSS, DeepSeek et Qwen.
TEMPERATURE = 0

# La tâche exige uniquement une phrase corrigée.
MAX_TOKENS = 500

# Durée de la temporisation après chaque phrase, en secondes.
REQUEST_DELAY = 2

# Nombre de tentatives supplémentaires en cas d’erreur ou de réponse vide.
MAX_RETRIES = 1

# Temporisation après une erreur 402 ou 429.
WAIT_INTERVAL = 20


# ============================================================
# 3. Fonctions utilitaires pour Excel
# ============================================================

def find_column(
    worksheet: Worksheet,
    column_name: str,
) -> int:
    """Recherche un nom de colonne dans la première ligne du classeur Excel.
    
    Renvoie son indice. En cas d’absence, affiche les colonnes disponibles
    et interrompt le programme."""

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
        f'找不到输入列 "{column_name}"。\n'
        f"当前 Excel 中的列名为：\n"
        f"{available_columns}"
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
# 4. Identification des erreurs API
# ============================================================

def is_credit_error(error_message: str) -> bool:
    """Identifie une erreur d’épuisement du quota Hugging Face."""

    message = str(error_message).lower()

    return (
        "error code: 402" in message
        or "status code: 402" in message
        or "depleted your monthly included credits" in message
        or "purchase pre-paid credits" in message
        or "payment required" in message
        or "insufficient balance" in message
    )


def is_rate_limit_error(error_message: str) -> bool:
    """Identifie une erreur de limitation du débit des requêtes."""

    message = str(error_message).lower()

    return (
        "error code: 429" in message
        or "status code: 429" in message
        or "rate limit" in message
        or "too many requests" in message
    )


def is_model_not_supported_error(
    error_message: str,
) -> bool:
    """Identifie une incompatibilité du modèle avec l’interface
    chat.completions de Hugging Face Router."""

    message = str(error_message).lower()

    return (
        "model_not_supported" in message
        or "is not a chat model" in message
        or "is not a chatmodel" in message
        or "model is not supported" in message
        or "no provider available" in message
        or "not supported by any provider" in message
        or "no inference provider" in message
    )


def is_access_error(error_message: str) -> bool:
    """Identifie une erreur liée aux conditions d’accès au modèle,
    aux permissions du jeton ou à un modèle à accès restreint."""

    message = str(error_message).lower()

    return (
        "gated repo" in message
        or "gated model" in message
        or "access to model" in message
        or "you are not authorized" in message
        or "unauthorized" in message
        or "forbidden" in message
        or "error code: 401" in message
        or "status code: 401" in message
        or "error code: 403" in message
        or "status code: 403" in message
        or "must be authenticated" in message
        or "accept the conditions" in message
        or "accept the license" in message
    )


# ============================================================
# 5. Nettoyage des sorties du modèle
# ============================================================

def clean_model_output(output: str) -> str:
    """Applique un nettoyage minimal à la sortie du modèle.
    
    Supprime les espaces périphériques, les blocs de code Markdown et les
    préfixes standard éventuels, sans reformuler la phrase française produite."""

    cleaned = output.strip()

    # Supprimer les délimiteurs de blocs de code Markdown.
    if cleaned.startswith("```") and cleaned.endswith("```"):
        cleaned = cleaned[3:-3].strip()

        if cleaned.lower().startswith("text"):
            cleaned = cleaned[4:].strip()

        elif cleaned.lower().startswith("french"):
            cleaned = cleaned[6:].strip()

        elif cleaned.lower().startswith("français"):
            cleaned = cleaned[8:].strip()

    # Supprimer les préfixes standard éventuellement ajoutés par le modèle.
    prefixes = [
        "Phrase corrigée :",
        "Phrase corrigée:",
        "Correction :",
        "Correction:",
        "Réponse :",
        "Réponse:",
    ]

    for prefix in prefixes:
        if cleaned.lower().startswith(prefix.lower()):
            cleaned = cleaned[len(prefix):].strip()
            break

    # Retirer uniquement les guillemets externes lorsqu’ils encadrent la phrase entière.
    quotation_pairs = [
        ('"', '"'),
        ("“", "”"),
        ("«", "»"),
    ]

    for opening, closing in quotation_pairs:
        if (
            cleaned.startswith(opening)
            and cleaned.endswith(closing)
            and len(cleaned) > 2
        ):
            cleaned = cleaned[
                len(opening):-len(closing)
            ].strip()
            break

    return cleaned


# ============================================================
# 6. Appel à Gemma 3 27B
# ============================================================

def call_model(
    client: OpenAI,
    system_prompt: str,
    sentence: str,
) -> dict[str, Any]:
    """Interroge Gemma 3 27B.
    
    Une réponse non vide est renvoyée directement. Les réponses vides et les
    erreurs API ordinaires autorisent MAX_RETRIES tentatives supplémentaires.
    Les erreurs 402 et 429 entraînent une nouvelle tentative après WAIT_INTERVAL
    secondes. Une interface non prise en charge ou une erreur d’autorisation
    provoque l’arrêt du traitement."""

    normal_attempts = 0
    total_api_attempts = 0
    credit_waits = 0
    rate_limit_waits = 0
    total_request_time = 0.0

    while True:
        total_api_attempts += 1
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

            if not response.choices:
                output = ""
            else:
                raw_output = (
                    response
                    .choices[0]
                    .message
                    .content
                )

                if raw_output is None:
                    output = ""
                else:
                    output = clean_model_output(
                        str(raw_output)
                    )

            # ------------------------------------------------
            # Réception d’une réponse non vide.
            # ------------------------------------------------

            if output:
                return {
                    "output": output,
                    "status": "OK",
                    "time": total_request_time,
                    "attempts": total_api_attempts,
                    "normal_retries": normal_attempts,
                    "credit_waits": credit_waits,
                    "rate_limit_waits": (
                        rate_limit_waits
                    ),
                }

            # ------------------------------------------------
            # Réponse vide du modèle.
            # ------------------------------------------------

            normal_attempts += 1

            print(
                f"模型第 {total_api_attempts} 次请求"
                "返回空输出。"
            )

            if normal_attempts > MAX_RETRIES:
                return {
                    "output": "",
                    "status": (
                        "FAILED: 模型返回空输出"
                    ),
                    "time": total_request_time,
                    "attempts": total_api_attempts,
                    "normal_retries": normal_attempts,
                    "credit_waits": credit_waits,
                    "rate_limit_waits": (
                        rate_limit_waits
                    ),
                }

            print("等待 3 秒后重试当前句……")
            time.sleep(3)

        except KeyboardInterrupt:
            print("\n收到 Ctrl+C。")
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
            # Interface non prise en charge par le modèle.
            # ------------------------------------------------

            if is_model_not_supported_error(
                error_text
            ):
                raise RuntimeError(
                    "\nGemma 3 27B 无法通过当前 "
                    "Hugging Face Router 的 "
                    "chat.completions 接口调用。\n\n"
                    f"模型：{MODEL_NAME}\n\n"
                    f"原始错误：\n{error_text}\n\n"
                    "程序已立即停止，避免整份语料"
                    "产生失败记录。"
                ) from error

            # ------------------------------------------------
            # Erreur d’autorisation d’accès au modèle.
            # ------------------------------------------------

            if is_access_error(error_text):
                raise RuntimeError(
                    "\n无法访问 Gemma 3 27B。\n\n"
                    "可能原因：\n"
                    "1. 尚未在 Hugging Face 模型页面"
                    "接受 Gemma 使用条款；\n"
                    "2. HF_TOKEN 没有读取权限；\n"
                    "3. HF_TOKEN 没有调用 "
                    "Inference Providers 的权限；\n"
                    "4. Token 已失效。\n\n"
                    f"模型：{MODEL_NAME}\n\n"
                    f"原始错误：\n{error_text}"
                ) from error

            # ------------------------------------------------
            # 402 : quota épuisé
            # ------------------------------------------------

            if is_credit_error(error_text):
                credit_waits += 1

                print("\n" + "!" * 70)
                print(
                    "检测到 Hugging Face 额度不足。"
                )
                print(
                    f"{WAIT_INTERVAL} 秒后重新请求"
                    "当前句。"
                )
                print(
                    f"额度等待次数：{credit_waits}"
                )
                print(
                    "程序不会跳到下一句。"
                )
                print(
                    "按 Ctrl+C 可以安全停止。"
                )
                print("!" * 70)

                time.sleep(WAIT_INTERVAL)
                continue

            # ------------------------------------------------
            # 429 : limitation du débit
            # ------------------------------------------------

            if is_rate_limit_error(error_text):
                rate_limit_waits += 1

                print("\n" + "!" * 70)
                print("检测到请求频率限制。")
                print(
                    f"{WAIT_INTERVAL} 秒后重新请求"
                    "当前句。"
                )
                print(
                    f"限流等待次数："
                    f"{rate_limit_waits}"
                )
                print(
                    "程序不会跳到下一句。"
                )
                print(
                    "按 Ctrl+C 可以安全停止。"
                )
                print("!" * 70)

                time.sleep(WAIT_INTERVAL)
                continue

            # ------------------------------------------------
            # Autres erreurs API
            # ------------------------------------------------

            normal_attempts += 1

            print(
                f"第 {total_api_attempts} 次请求失败："
            )
            print(error_text)

            if normal_attempts > MAX_RETRIES:
                return {
                    "output": "",
                    "status": (
                        f"FAILED: {error_text}"
                    ),
                    "time": total_request_time,
                    "attempts": total_api_attempts,
                    "normal_retries": normal_attempts,
                    "credit_waits": credit_waits,
                    "rate_limit_waits": (
                        rate_limit_waits
                    ),
                }

            print("等待 3 秒后重试当前句……")
            time.sleep(3)


# ============================================================
# 7. Point d’entrée
# ============================================================

def main() -> None:
    """Lit le corpus, interroge Gemma 3 27B et enregistre les résultats
    phrase par phrase dans un classeur Excel."""

    # --------------------------------------------------------
    # Chargement des variables d’environnement.
    # --------------------------------------------------------

    load_dotenv(PROJECT_ROOT / ".env")

    hf_token = os.getenv("HF_TOKEN")

    if not hf_token:
        raise RuntimeError(
            "没有读取到 HF_TOKEN。\n\n"
            "请检查项目根目录中的 .env 文件。\n\n"
            "正确格式示例：\n"
            "HF_TOKEN=YOUR_HF_TOKEN"
        )

    # --------------------------------------------------------
    # Vérification des fichiers.
    # --------------------------------------------------------

    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"找不到输入文件：\n"
            f"{INPUT_FILE}\n\n"
            "请确认文件名为 corpus_107.xlsx。"
        )

    if not PROMPT_FILE.exists():
        raise FileNotFoundError(
            f"找不到 Prompt 文件：\n"
            f"{PROMPT_FILE}"
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
            f"Prompt 文件为空：\n{PROMPT_FILE}"
        )

    # --------------------------------------------------------
    # Initialisation du client Hugging Face Router.
    # --------------------------------------------------------

    client = OpenAI(
        base_url=(
            "https://router.huggingface.co/v1"
        ),
        api_key=hf_token,
    )

    # --------------------------------------------------------
    # Lecture des résultats existants ou du corpus source.
    # --------------------------------------------------------

    if OUTPUT_FILE.exists():
        print(
            "检测到已有 Gemma 3 结果文件。"
        )
        print(
            "程序将从上次中断的位置继续运行。"
        )
        print("读取文件：", OUTPUT_FILE)

        workbook = load_workbook(
            OUTPUT_FILE
        )

    else:
        print(
            "没有检测到旧的 Gemma 3 结果文件。"
        )
        print(
            "程序将开始新的 Gemma 3 实验。"
        )
        print("读取原始语料：", INPUT_FILE)

        workbook = load_workbook(
            INPUT_FILE
        )

    worksheet = workbook.active

    # --------------------------------------------------------
    # Identification de la colonne d’entrée.
    # --------------------------------------------------------

    input_column = find_column(
        worksheet,
        INPUT_COLUMN_NAME,
    )

    # --------------------------------------------------------
    # Création des colonnes de résultats Gemma.
    # --------------------------------------------------------

    output_column = get_or_create_column(
        worksheet,
        f"Sortie_{MODEL_SHORT_NAME}",
    )

    status_column = get_or_create_column(
        worksheet,
        f"Statut_{MODEL_SHORT_NAME}",
    )

    time_column = get_or_create_column(
        worksheet,
        f"Temps_secondes_{MODEL_SHORT_NAME}",
    )

    attempts_column = get_or_create_column(
        worksheet,
        f"Tentatives_{MODEL_SHORT_NAME}",
    )

    normal_retries_column = get_or_create_column(
        worksheet,
        f"Relances_normales_{MODEL_SHORT_NAME}",
    )

    credit_waits_column = get_or_create_column(
        worksheet,
        f"Attentes_credit_{MODEL_SHORT_NAME}",
    )

    rate_limit_waits_column = (
        get_or_create_column(
            worksheet,
            f"Attentes_limite_{MODEL_SHORT_NAME}",
        )
    )

    model_column = get_or_create_column(
        worksheet,
        f"Modele_{MODEL_SHORT_NAME}",
    )

    # --------------------------------------------------------
    # Initialisation des statistiques expérimentales.
    # --------------------------------------------------------

    total_rows = worksheet.max_row - 1

    processed_count = 0
    skipped_count = 0
    failed_count = 0

    print("\n" + "=" * 70)
    print("当前实验模型：", MODEL_NAME)
    print("语料总数：", total_rows)
    print("输入文件：", INPUT_FILE)
    print("Prompt 文件：", PROMPT_FILE)
    print("输出文件：", OUTPUT_FILE)
    print("Temperature：", TEMPERATURE)
    print("Max tokens：", MAX_TOKENS)
    print("=" * 70)

    # --------------------------------------------------------
    # Parcours des phrases du corpus.
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

            sentence = str(
                sentence_value
            ).strip()

            if not sentence:
                skipped_count += 1
                continue

            # ------------------------------------------------
            # Reprise du traitement : ignorer les résultats déjà validés par le statut OK.
            # ------------------------------------------------

            existing_output = worksheet.cell(
                row=row_number,
                column=output_column,
            ).value

            existing_status = worksheet.cell(
                row=row_number,
                column=status_column,
            ).value

            if (
                existing_output is not None
                and str(existing_output).strip()
                and existing_status == "OK"
            ):
                print(
                    f"第 {row_number} 行已有 "
                    "Gemma 3 成功结果，跳过。"
                )

                skipped_count += 1
                continue

            current_sentence_number = (
                row_number - 1
            )

            print("\n" + "=" * 70)
            print(
                f"正在处理第 "
                f"{current_sentence_number} / "
                f"{total_rows} 句"
            )
            print("Excel 行号：", row_number)
            print("模型：", MODEL_NAME)
            print("输入：", sentence)

            # ------------------------------------------------
            # Appel au modèle.
            # ------------------------------------------------

            result = call_model(
                client=client,
                system_prompt=system_prompt,
                sentence=sentence,
            )

            output = result["output"]
            status = result["status"]
            elapsed_time = result["time"]
            attempts = result["attempts"]
            normal_retries = result[
                "normal_retries"
            ]
            credit_waits = result[
                "credit_waits"
            ]
            rate_limit_waits = result[
                "rate_limit_waits"
            ]

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
                value=round(
                    elapsed_time,
                    3,
                ),
            )

            worksheet.cell(
                row=row_number,
                column=attempts_column,
                value=attempts,
            )

            worksheet.cell(
                row=row_number,
                column=normal_retries_column,
                value=normal_retries,
            )

            worksheet.cell(
                row=row_number,
                column=credit_waits_column,
                value=credit_waits,
            )

            worksheet.cell(
                row=row_number,
                column=rate_limit_waits_column,
                value=rate_limit_waits,
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
            print("API 请求次数：", attempts)
            print(
                "普通重试次数：",
                normal_retries,
            )
            print(
                "额度等待次数：",
                credit_waits,
            )
            print(
                "限流等待次数：",
                rate_limit_waits,
            )
            print(
                f"API 累计耗时："
                f"{elapsed_time:.2f} 秒"
            )

            # ------------------------------------------------
            # Enregistrer le classeur après chaque phrase.
            # ------------------------------------------------

            workbook.save(OUTPUT_FILE)

            print("已保存：", OUTPUT_FILE)

            time.sleep(REQUEST_DELAY)

    except KeyboardInterrupt:
        print("\n" + "=" * 70)
        print(
            "用户通过 Ctrl+C 停止程序。"
        )
        print(
            "正在保存 Gemma 3 当前结果……"
        )

        workbook.save(OUTPUT_FILE)

        print("结果已保存：", OUTPUT_FILE)
        print(
            "下次运行同一程序时，"
            "将继续处理尚未成功的句子。"
        )
        print("=" * 70)

        return

    except RuntimeError:
        print("\n" + "=" * 70)
        print(
            "程序遇到不可继续的错误。"
        )
        print(
            "正在保存 Gemma 3 当前结果……"
        )

        workbook.save(OUTPUT_FILE)

        print(
            "当前结果已保存：",
            OUTPUT_FILE,
        )
        print("=" * 70)

        raise

    except Exception:
        print("\n" + "=" * 70)
        print(
            "程序出现未预期错误。"
        )
        print(
            "正在保存当前已完成的结果……"
        )

        workbook.save(OUTPUT_FILE)

        print(
            "当前结果已保存：",
            OUTPUT_FILE,
        )
        print("=" * 70)

        raise

    # --------------------------------------------------------
    # Bilan du traitement.
    # --------------------------------------------------------

    workbook.save(OUTPUT_FILE)

    print("\n" + "=" * 70)
    print("Gemma 3 27B 实验运行结束。")
    print("本次新处理：", processed_count)
    print(
        "跳过已有成功结果：",
        skipped_count,
    )
    print(
        "失败或空输出：",
        failed_count,
    )
    print("结果文件：", OUTPUT_FILE)
    print("=" * 70)


if __name__ == "__main__":
    main()