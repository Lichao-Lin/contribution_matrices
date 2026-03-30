import argparse
import math
import re
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook, load_workbook


DEFAULT_INPUT = Path(
    r"C:\Users\megamind\Desktop\proquest全文清洗2025.12.16 的定稿(2).xlsm"
)
DEFAULT_OUTPUT = Path(
    r"C:\Users\megamind\Desktop\tf-idf-output.xlsx"
)
DEFAULT_SHEET = "Sheet1"
DEFAULT_COLUMN = 3

TERMS = [
    "thailand",
    "cooperation",
    "tesla",
    "european",
    "tariff",
    "trade",
    "country",
    "international",
    "export",
    "car",
    "energy",
    "development",
    "byd",
    "industry",
    "china",
    "global",
    "price",
    "europe",
    "president",
    "economic",
    "the eu",
    "foreign",
    "support",
    "share",
    "ministry",
    "work",
    "company",
    "government",
    "market",
    "need",
    "new",
    "vehicle",
    "electric",
    "chinese",
    "unfair",
    "technological",
    "cheap",
    "strategic",
    "mutual",
    "local",
    "sustainable",
    "japanese",
    "strong",
    "massive",
]

TOKEN_RE = re.compile(r"[a-z]+(?:'[a-z]+)?")


@dataclass
class TermStats:
    term: str
    count: int
    document_count: int
    tf_idf: float


def normalize_text(value) -> str:
    return str(value or "").strip().lower()


def build_pattern(term: str) -> re.Pattern[str]:
    return re.compile(rf"(?<![a-z]){re.escape(term.lower())}(?![a-z])")


def read_documents(file_path: Path, sheet_name: str, column_index: int) -> list[str]:
    workbook = load_workbook(file_path, read_only=True, data_only=True)
    worksheet = workbook[sheet_name]
    documents = []

    for row in worksheet.iter_rows(min_row=2, values_only=True):
        if len(row) < column_index:
            continue
        text = normalize_text(row[column_index - 1])
        if text:
            documents.append(text)

    return documents


def count_total_tokens(documents: list[str]) -> int:
    return sum(len(TOKEN_RE.findall(document)) for document in documents)


def calculate_tfidf(documents: list[str], terms: list[str]) -> list[TermStats]:
    total_documents = len(documents)
    total_tokens = count_total_tokens(documents)

    if total_documents == 0:
        raise ValueError("未读取到任何文档内容，请检查 sheet 名称或列号。")

    if total_tokens == 0:
        raise ValueError("第三列没有可用于分词的英文 token。")

    results = []
    for term in terms:
        pattern = build_pattern(term)
        count = sum(len(pattern.findall(document)) for document in documents)
        document_count = sum(1 for document in documents if pattern.search(document))

        if count == 0 or document_count == 0:
            tf_idf = 0.0
        else:
            tf = count / total_tokens
            idf = math.log(total_documents / document_count)
            tf_idf = tf * idf

        results.append(
            TermStats(
                term=term,
                count=count,
                document_count=document_count,
                tf_idf=tf_idf,
            )
        )

    return results


def write_output(output_path: Path, results: list[TermStats]) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "sheet1"
    worksheet.append(["单词", "次数", "条数", "tf-idf"])

    for item in results:
        worksheet.append([item.term, item.count, item.document_count, item.tf_idf])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="统计指定词在 Excel 第三列中的 corpus-level tf-idf")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT, help="输入 Excel 文件路径")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT, help="输出 Excel 文件路径")
    parser.add_argument("--sheet", default=DEFAULT_SHEET, help="工作表名称")
    parser.add_argument("--column", type=int, default=DEFAULT_COLUMN, help="目标列号，从 1 开始")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    documents = read_documents(args.input, args.sheet, args.column)
    results = calculate_tfidf(documents, TERMS)
    write_output(args.output, results)
    print(f"已输出: {args.output}")


if __name__ == "__main__":
    main()
