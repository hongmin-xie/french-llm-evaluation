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

OUTPUT_FILE = (
    PROJECT_ROOT
    / "result"
    / "results_107_qwen3_32b.xlsx"
)

PROMPT_FILE = (
    PROJECT_ROOT
    / "prompts"
    / "prompt_v1.txt"
)


# ============================================================
# 2. Configuration du modèle et paramètres expérimentaux
# ============================================================

MODEL_NAME = "Qwen/Qwen3-32B"

MODEL_SHORT_NAME = "qwen3_32b"

# Nom de la colonne du classeur contenant les phrases erronées.
INPUT_COLUMN_NAME = "Phrase erronée"

# Conserver des paramètres identiques pour permettre la comparaison des modèles.
TEMPERATURE = 0

# Une limite de 500 tokens est suffisante pour la production d’une seule phrase.
MAX_TOKENS = 500

# Temporisation de deux secondes après chaque phrase.
REQUEST_DELAY = 2

# Une tentative supplémentaire en cas de réponse vide ou d’erreur API.
MAX_RETRIES = 1

# En cas d’erreur 402 ou 429, réessayer toutes les 20 secondes.
WAIT_INTERVAL = 20


# ============================================================
# 3. Fonctions utilitaires pour Excel
# ============================================================

def find_column(
    worksheet: Worksheet,
    column_name: str,
) -> int:
    """Recherche un nom de colonne dans la première ligne du classeur Excel.
    
    Renvoie son indice ou affiche les colonnes disponibles si elle est absente."""

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
        f"当前 Excel 中的列名为：{available_columns}"
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
        or "depleted your monthly included credits" in message
        or "purchase pre-paid credits" in message
        or "payment required" in message
    )


def is_rate_limit_error(error_message: str) -> bool:
    """Identifie une limitation du débit des requêtes."""

    message = str(error_message).lower()

    return (
        "error code: 429" in message
        or "rate limit" in message
        or "too many requests" in message
    )


def is_model_not_supported_error(
    error_message: str,
) -> bool:
    """Identifie une incompatibilité du modèle avec l’interface chat.completions."""

    message = str(error_message).lower()

    return (
        "model_not_supported" in message
        or "is not a chat model" in message
        or "is not a chatmodel" in message
        or "model is not supported" in message
        or "no provider available" in message
        or "not supported by any provider" in message
    )


def is_invalid_parameter_error(
    error_message: str,
) -> bool:
    """Identifie un paramètre supplémentaire non pris en charge par le fournisseur.
    
    Certains fournisseurs peuvent refuser des paramètres transmis via extra_body."""

    message = str(error_message).lower()

    return (
        "extra_body" in message
        or "chat_template_kwargs" in message
        or "enable_thinking" in message
        or "unexpected keyword" in message
        or "unknown parameter" in message
    )


# ============================================================
# 5. Nettoyage des sorties du modèle
# ============================================================

def clean_model_output(output: str) -> str:
    """Applique un nettoyage minimal à la sortie de Qwen.
    
    Supprime les espaces périphériques et les blocs de raisonnement délimités
    par les balises <think> et </think>, sans reformuler le contenu de la phrase."""

    cleaned = output.strip()

    # Certains fournisseurs Qwen peuvent retourner des balises de raisonnement.
    if "</think>" in cleaned:
        cleaned = cleaned.split("</think>", 1)[1].strip()

    return cleaned


# ============================================================
# 6. Appel à Qwen3-32B
# ============================================================

def call_model(
    client: OpenAI,
    system_prompt: str,
    sentence: str,
) -> dict[str, Any]:
    """Interroge Qwen3-32B.
    
    Une réponse non vide est renvoyée directement. Les réponses vides et les
    erreurs API ordinaires autorisent au plus MAX_RETRIES tentatives supplémentaires.
    Les erreurs 402 et 429 entraînent une nouvelle tentative après 20 secondes.
    Une interface chat.completions non prise en charge provoque l’arrêt. Si le
    fournisseur refuse le paramètre de désactivation du raisonnement, une requête
    standard est utilisée."""

    normal_attempts = 0
    credit_waits = 0
    rate_limit_waits = 0
    total_request_time = 0.0

    # Tenter en priorité de désactiver le mode de raisonnement de Qwen3.
    use_thinking_parameter = True

    while True:
        start_time = time.perf_counter()

        try:
            request_arguments: dict[str, Any] = {
                "model": MODEL_NAME,
                "messages": [
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
                "temperature": TEMPERATURE,
                "max_tokens": MAX_TOKENS,
            }

            # Qwen3 prend en charge les modes avec et sans raisonnement.
            # L’expérience de correction grammaticale requiert uniquement la phrase finale.
            if use_thinking_parameter:
                request_arguments["extra_body"] = {
                    "chat_template_kwargs": {
                        "enable_thinking": False
                    }
                }

            response = client.chat.completions.create(
                **request_arguments
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
                output = clean_model_output(output)

            if output:
                return {
                    "output": output,
                    "status": "OK",
                    "time": total_request_time,
                    "attempts": normal_attempts + 1,
                    "credit_waits": credit_waits,
                    "rate_limit_waits": rate_limit_waits,
                    "thinking_disabled": (
                        use_thinking_parameter
                    ),
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
                    "rate_limit_waits": rate_limit_waits,
                    "thinking_disabled": (
                        use_thinking_parameter
                    ),
                }

            print("等待 3 秒后重试当前句……")
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
            # Interface chat.completions non prise en charge : arrêt du traitement.
            # ------------------------------------------------

            if is_model_not_supported_error(error_text):
                raise RuntimeError(
                    "\n当前模型无法通过 Hugging Face Router 的 "
                    "chat.completions 接口调用。\n\n"
                    f"模型：{MODEL_NAME}\n\n"
                    f"原始错误：\n{error_text}\n\n"
                    "程序已立即停止，避免整份语料全部失败。"
                ) from error

            # ------------------------------------------------
            # Paramètre de désactivation du raisonnement non pris en charge par le fournisseur.
            # ------------------------------------------------

            if (
                use_thinking_parameter
                and is_invalid_parameter_error(error_text)
            ):
                print("\n" + "!" * 70)
                print(
                    "当前 Provider 不支持 "
                    "enable_thinking=False 参数。"
                )
                print(
                    "程序将自动移除该参数，"
                    "并重新请求当前句。"
                )
                print("!" * 70)

                use_thinking_parameter = False
                continue

            # ------------------------------------------------
            # 402 : quota épuisé
            # ------------------------------------------------

            if is_credit_error(error_text):
                credit_waits += 1

                print("\n" + "!" * 70)
                print("检测到 Hugging Face 额度不足。")
                print(
                    f"{WAIT_INTERVAL} 秒后重新请求当前句。"
                )
                print(
                    f"额度等待次数：{credit_waits}"
                )
                print("按 Ctrl+C 可以停止程序。")
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
                    f"{WAIT_INTERVAL} 秒后重新请求当前句。"
                )
                print(
                    f"限流等待次数：{rate_limit_waits}"
                )
                print("按 Ctrl+C 可以停止程序。")
                print("!" * 70)

                time.sleep(WAIT_INTERVAL)
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
                    "rate_limit_waits": rate_limit_waits,
                    "thinking_disabled": (
                        use_thinking_parameter
                    ),
                }

            print("等待 3 秒后重试当前句……")
            time.sleep(3)


# ============================================================
# 7. Point d’entrée
# ============================================================

def main() -> None:
    """Traite par lots les phrases du classeur Excel."""

    # --------------------------------------------------------
    # Chargement de HF_TOKEN depuis .env.
    # --------------------------------------------------------

    load_dotenv(PROJECT_ROOT / ".env")

    hf_token = os.getenv("HF_TOKEN")

    if not hf_token:
        raise RuntimeError(
            "没有读取到 HF_TOKEN。\n"
            "请检查项目根目录下的 .env 文件。\n\n"
            "正确格式示例：\n"
            "HF_TOKEN=YOUR_HF_TOKEN"
        )

    # --------------------------------------------------------
    # Vérification du corpus et du fichier de prompt.
    # --------------------------------------------------------

    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"找不到输入文件：\n{INPUT_FILE}\n\n"
            "请确认文件名为 corpus_107.xlsx。"
        )

    if not PROMPT_FILE.exists():
        raise FileNotFoundError(
            f"找不到 Prompt 文件：\n{PROMPT_FILE}"
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
        base_url="https://router.huggingface.co/v1",
        api_key=hf_token,
    )

    # --------------------------------------------------------
    # Reprise des résultats Qwen existants ou création du classeur depuis le corpus.
    # --------------------------------------------------------

    if OUTPUT_FILE.exists():
        print("检测到已有 Qwen3 结果文件。")
        print("程序将继续上一次 Qwen3 实验。")
        print("读取文件：", OUTPUT_FILE)

        workbook = load_workbook(OUTPUT_FILE)

    else:
        print("没有检测到旧的 Qwen3 结果文件。")
        print("程序将开始新的 Qwen3 实验。")
        print("读取原始语料：", INPUT_FILE)

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
    # Création des colonnes de résultats Qwen3.
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

    credit_waits_column = get_or_create_column(
        worksheet,
        f"Attentes_credit_{MODEL_SHORT_NAME}",
    )

    rate_limit_waits_column = get_or_create_column(
        worksheet,
        f"Attentes_limite_{MODEL_SHORT_NAME}",
    )

    model_column = get_or_create_column(
        worksheet,
        f"Modele_{MODEL_SHORT_NAME}",
    )

    thinking_column = get_or_create_column(
        worksheet,
        f"Thinking_desactive_{MODEL_SHORT_NAME}",
    )

    # --------------------------------------------------------
    # Initialisation des compteurs.
    # --------------------------------------------------------

    total_rows = worksheet.max_row - 1

    processed_count = 0
    skipped_count = 0
    failed_count = 0

    print("\n" + "=" * 70)
    print("当前实验模型：", MODEL_NAME)
    print("语料行数：", total_rows)
    print("输入文件：", INPUT_FILE)
    print("Prompt 文件：", PROMPT_FILE)
    print("输出文件：", OUTPUT_FILE)
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

            # Vérifier si la phrase dispose déjà d’un résultat valide.
            existing_output = worksheet.cell(
                row=row_number,
                column=output_column,
            ).value

            existing_status = worksheet.cell(
                row=row_number,
                column=status_column,
            ).value

            # Ignorer uniquement les lignes dont la sortie est renseignée et le statut est OK.
            if (
                existing_output
                and existing_status == "OK"
            ):
                print(
                    f"第 {row_number} 行已有 Qwen3 "
                    "成功结果，跳过。"
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
            print("Excel 行号：", row_number)
            print("模型：", MODEL_NAME)
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
            rate_limit_waits = result[
                "rate_limit_waits"
            ]
            thinking_disabled = result[
                "thinking_disabled"
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
                column=rate_limit_waits_column,
                value=rate_limit_waits,
            )

            worksheet.cell(
                row=row_number,
                column=model_column,
                value=MODEL_NAME,
            )

            worksheet.cell(
                row=row_number,
                column=thinking_column,
                value=thinking_disabled,
            )

            processed_count += 1

            if status != "OK":
                failed_count += 1

            # ------------------------------------------------
            # Affichage du résultat courant dans le terminal.
            # ------------------------------------------------

            print("输出：", output)
            print("状态：", status)
            print("普通尝试次数：", attempts)
            print("额度等待次数：", credit_waits)
            print(
                "限流等待次数：",
                rate_limit_waits,
            )
            print(
                "Thinking 已关闭：",
                thinking_disabled,
            )
            print(
                f"API 累计耗时："
                f"{elapsed_time:.2f} 秒"
            )

            # Enregistrer le classeur après chaque phrase.
            workbook.save(OUTPUT_FILE)

            print("已保存：", OUTPUT_FILE)

            # Temporisation avant la phrase suivante.
            time.sleep(REQUEST_DELAY)

    except KeyboardInterrupt:
        print("\n" + "=" * 70)
        print("用户通过 Ctrl+C 停止程序。")
        print("正在保存 Qwen3 当前结果……")

        workbook.save(OUTPUT_FILE)

        print("结果已保存：", OUTPUT_FILE)
        print(
            "下次运行同一程序时，"
            "将继续处理尚未成功的句子。"
        )
        print("=" * 70)

        return

    except RuntimeError:
        # Interface non prise en charge ou autre erreur nécessitant un arrêt.
        print("\n" + "=" * 70)
        print("程序遇到不可继续的错误。")
        print("正在保存当前 Qwen3 结果……")

        workbook.save(OUTPUT_FILE)

        print("当前结果已保存：", OUTPUT_FILE)
        print("=" * 70)

        raise

    # --------------------------------------------------------
    # Bilan du traitement.
    # --------------------------------------------------------

    workbook.save(OUTPUT_FILE)

    print("\n" + "=" * 70)
    print("Qwen3-32B 实验运行结束。")
    print("本次新处理：", processed_count)
    print("跳过已有成功结果：", skipped_count)
    print("失败或空输出：", failed_count)
    print("结果文件：", OUTPUT_FILE)
    print("=" * 70)


if __name__ == "__main__":
    main()