
import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
import time
import shutil
from pathlib import Path
from tqdm import tqdm

# =========================================================
# 1. 기본 설정 및 LibreOffice UNO 환경 구성
# =========================================================

BASE_DIR = Path(__file__).resolve().parent
INPUT_DIR = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"
SPREADSHEET_EXTENSIONS = {".xlsx", ".xls", ".xlsm", ".ods"}
SUPPORTED_EXTENSIONS = SPREADSHEET_EXTENSIONS | {".docx"}

LO_PROGRAM = Path(
    r"C:\Program Files\LibreOffice\program"
)

UNO_HOST = "127.0.0.1"
UNO_PORT = 2002

UNO_CONNECTION = (
    f"uno:socket,host={UNO_HOST},port={UNO_PORT};"
    "urp;StarOffice.ComponentContext"
)

# LibreOffice UNO 모듈 경로 설정
if LO_PROGRAM.exists():
    if hasattr(os, "add_dll_directory"):
        os.add_dll_directory(str(LO_PROGRAM))

    os.environ["PATH"] = (
        str(LO_PROGRAM) + os.pathsep +
        os.environ.get("PATH", "")
    )

    sys.path.insert(0, str(LO_PROGRAM))

import uno
from com.sun.star.beans import PropertyValue

# =========================================================
# 2. LibreOffice 실행 파일 찾기
# =========================================================

def find_libreoffice():

    candidates = [
        LO_PROGRAM / "soffice.exe",
        Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
    ]

    for path in candidates:
        if path.exists():
            return str(path)

    raise FileNotFoundError(
        "LibreOffice soffice.exe를 찾을 수 없습니다."
    )


# =========================================================
# 3. UNO Property 생성
# =========================================================

def make_property(name, value):

    prop = PropertyValue()
    prop.Name = name
    prop.Value = value

    return prop


# =========================================================
# 4. LibreOffice 서버 실행 및 연결
# =========================================================

def start_libreoffice():

    soffice_path = find_libreoffice()

    profile_dir = Path(
        tempfile.mkdtemp(prefix="lo_markdown_profile_")
    )

    profile_url = profile_dir.as_uri()

    log_path = BASE_DIR / "libreoffice_error.log"

    log_file = open(
        log_path,
        "w",
        encoding="utf-8"
    )

    command = [
        soffice_path,
        "--headless",
        "--nologo",
        "--nodefault",
        "--norestore",
        "--nofirststartwizard",
        f"-env:UserInstallation={profile_url}",
        (
            f"--accept=socket,host={UNO_HOST},"
            f"port={UNO_PORT},tcpNoDelay=1;"
            "urp;StarOffice.ServiceManager"
        ),
    ]

    print("\n[1/4] LibreOffice 서버 실행 중...")

    process = subprocess.Popen(
        command,
        stdout=log_file,
        stderr=subprocess.STDOUT
    )

    local_context = uno.getComponentContext()

    resolver = (
        local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver",
            local_context
        )
    )

    print("[2/4] UNO 서버 연결 대기 중...")

    last_error = None
    try:

        for attempt in tqdm(
            range(60),
            desc="UNO 연결",
            unit="초"
        ):

            try:

                context = resolver.resolve(
                    UNO_CONNECTION
                )

                desktop = (
                    context.ServiceManager.createInstanceWithContext(
                        "com.sun.star.frame.Desktop",
                        context
                    )
                )

                print("\nUNO 서버 연결 성공!")

                return process, desktop, profile_dir, log_file

            except Exception as error:

                last_error = error

                if process.poll() is not None:

                    raise RuntimeError(
                        f"LibreOffice 프로세스 종료. "
                        f"로그 확인: {log_path}"
                    )

                time.sleep(1)

        raise TimeoutError(
            f"UNO 서버 연결 시간이 초과되었습니다. "
            f"로그 확인: {log_path}. 마지막 오류: {last_error}"
        ) from last_error

    except Exception:

        if process.poll() is None:
            process.terminate()

        process.wait(timeout=10)

        log_file.close()

        shutil.rmtree(
            profile_dir,
            ignore_errors=True
        )

        raise


# =========================================================
# 5. LibreOffice 문서 열기
# =========================================================

def open_document(desktop, file_path):

    file_url = uno.systemPathToFileUrl(
        str(Path(file_path).resolve())
    )

    properties = (
        make_property("Hidden", True),
        make_property("ReadOnly", True),
        make_property("UpdateDocMode", 3),
        make_property("MacroExecutionMode", 0),
    )

    document = desktop.loadComponentFromURL(
        file_url,
        "_blank",
        0,
        properties
    )

    if document is None:

        raise RuntimeError(
            f"문서를 열지 못했습니다: {file_path}"
        )

    return document


# =========================================================
# 6. 시트 데이터 추출
# =========================================================

def extract_sheet(sheet):

    cursor = sheet.createCursor()

    cursor.gotoEndOfUsedArea(True)

    address = cursor.RangeAddress

    start_row = address.StartRow
    end_row = address.EndRow

    start_col = address.StartColumn
    end_col = address.EndColumn

    grid = []

    total_rows = end_row - start_row + 1

    for row_index in range(start_row, end_row + 1):

        row_values = []

        for col_index in range(start_col, end_col + 1):

            cell = sheet.getCellByPosition(
                col_index,
                row_index
            )

            value = cell.getString()

            row_values.append(value.strip())

        grid.append(row_values)

    return {
        "grid": grid,
        "start_row": start_row + 1,
        "end_row": end_row + 1,
        "start_col": start_col + 1,
        "end_col": end_col + 1,
    }


# =========================================================
# 7. 셀 주소 변환
# =========================================================

def column_name(number):

    result = ""

    while number > 0:

        number, remainder = divmod(number - 1, 26)

        result = chr(65 + remainder) + result

    return result


def cell_address(row, col):

    return f"{column_name(col)}{row}"


def range_address(r1, c1, r2, c2):

    return (
        f"{cell_address(r1, c1)}:"
        f"{cell_address(r2, c2)}"
    )


# =========================================================
# 8. Markdown 특수문자 처리
# =========================================================

def escape_markdown(value):

    value = str(value)

    value = value.replace("\\", "\\\\")
    value = value.replace("|", "\\|")
    value = value.replace("\r\n", "<br>")
    value = value.replace("\n", "<br>")
    value = value.replace("\r", "<br>")

    return value


def safe_filename(name):

    name = re.sub(
        r'[\\/:*?"<>|]+',
        "_",
        name
    )

    return name.strip(" .") or "sheet"


# =========================================================
# 9. 표 및 텍스트 영역 판별
# =========================================================

def find_blocks(grid):

    row_info = []

    for row in grid:

        populated = [
            i for i, value in enumerate(row)
            if str(value).strip()
        ]

        row_info.append({
            "populated": populated,
            "candidate": len(populated) >= 2,
        })

    blocks = []

    i = 0

    while i < len(grid):

        info = row_info[i]

        if not info["populated"]:
            i += 1
            continue

        if info["candidate"]:

            j = i

            while (
                j < len(grid)
                and row_info[j]["candidate"]
            ):
                j += 1

            if j - i >= 2:

                table_rows = grid[i:j]

                used_cols = sorted({
                    col
                    for row in table_rows
                    for col in range(len(row))
                    if str(row[col]).strip()
                })

                if used_cols:

                    first_col = min(used_cols)
                    last_col = max(used_cols)

                    normalized = [
                        row[first_col:last_col + 1]
                        for row in table_rows
                    ]

                    blocks.append({
                        "type": "table",
                        "rows": normalized,
                        "range": range_address(
                            i + 1,
                            first_col + 1,
                            j,
                            last_col + 1
                        )
                    })

                    i = j
                    continue

        text_lines = []
        refs = []

        j = i

        while j < len(grid):

            if not row_info[j]["populated"]:
                break

            if row_info[j]["candidate"]:

                k = j

                while (
                    k < len(grid)
                    and row_info[k]["candidate"]
                ):
                    k += 1

                if k - j >= 2:
                    break

            populated = row_info[j]["populated"]

            values = []

            for col in populated:

                value = str(grid[j][col]).strip()

                if value:

                    values.append(value)

                    refs.append(
                        cell_address(j + 1, col + 1)
                    )

            if values:
                text_lines.append(" ".join(values))

            j += 1

        if text_lines:

            blocks.append({
                "type": "text",
                "text": "\n".join(text_lines),
                "range": (
                    refs[0]
                    if len(refs) == 1
                    else f"{refs[0]} ... {refs[-1]}"
                )
            })

        i = max(j, i + 1)

    return blocks


# =========================================================
# 10. Markdown 변환
# =========================================================

def blocks_to_markdown(sheet_name, blocks):

    lines = [
        f"# {sheet_name}",
        ""
    ]

    for block in blocks:

        if block["type"] == "text":

            lines.extend([
                block["text"],
                "",
                f"<!-- source: {block['range']} -->",
                ""
            ])

        elif block["type"] == "table":

            rows = block["rows"]

            if not rows:
                continue

            width = max(len(row) for row in rows)

            normalized = [
                row + [""] * (width - len(row))
                for row in rows
            ]

            header = normalized[0]

            header = [
                value if str(value).strip()
                else f"Column {i + 1}"
                for i, value in enumerate(header)
            ]

            lines.append(
                "| " + " | ".join(
                    escape_markdown(v)
                    for v in header
                ) + " |"
            )

            lines.append(
                "| " + " | ".join(
                    "---" for _ in header
                ) + " |"
            )

            for row in normalized[1:]:

                lines.append(
                    "| " + " | ".join(
                        escape_markdown(v)
                        for v in row
                    ) + " |"
                )

            lines.extend([
                "",
                f"<!-- source: {block['range']} -->",
                ""
            ])

    return "\n".join(lines).rstrip() + "\n"


# =========================================================
# 11. 파일명 중복 방지
# =========================================================

def unique_filename(index, sheet_name):

    name = safe_filename(sheet_name)

    return f"sheet_{index:02d}_{name}.md"


def extract_writer_table(table):
    """Read existing cells, including merged/split cells, without images."""
    cells = {}
    for name in table.getCellNames():
        match = re.fullmatch(r"([A-Z]+)([0-9]+)(?:\.[0-9]+)*", name)
        if match is None:
            raise ValueError(f"지원하지 않는 Word 표 셀 주소: {name}")
        col = 0
        for letter in match[1]:
            col = col * 26 + ord(letter) - ord("A") + 1
        row = int(match[2])
        value = table.getCellByName(name).getString().strip()
        # Split cells share a base address. Keep all their text in that cell.
        cells.setdefault((row, col), []).append(value)

    if not cells:
        return []
    return [
        [
            "\n".join(cells.get((row, col), []))
            for col in range(1, max(c for _, c in cells) + 1)
        ]
        for row in range(1, max(r for r, _ in cells) + 1)
    ]


def extract_writer_blocks(document):
    """Enumerate paragraphs and tables in document order."""
    blocks = []
    elements = document.getText().createEnumeration()
    paragraph_index = 0
    table_index = 0
    while elements.hasMoreElements():
        element = elements.nextElement()
        if element.supportsService("com.sun.star.text.TextTable"):
            table_index += 1
            blocks.append({
                "type": "table",
                "rows": extract_writer_table(element),
                "range": f"table:{table_index}",
            })
        elif element.supportsService("com.sun.star.text.Paragraph"):
            paragraph_index += 1
            text = element.getString().strip()
            if text:
                level = element.getPropertyValue("OutlineLevel")
                if level:
                    text = "#" * min(int(level), 6) + " " + text
                blocks.append({
                    "type": "text",
                    "text": text,
                    "range": f"paragraph:{paragraph_index}",
                })
    return blocks


def write_writer_output(document, input_file):
    blocks = extract_writer_blocks(document)
    filename = safe_filename(input_file.stem) + ".md"
    (OUTPUT_DIR / filename).write_text(
        blocks_to_markdown(input_file.stem, blocks), encoding="utf-8"
    )
    (OUTPUT_DIR / "manifest.json").write_text(
        json.dumps({
            "source_file": input_file.name,
            "engine": "LibreOffice Writer UNO",
            "file": filename,
            "blocks": blocks,
        }, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print("\n변환 완료!")
    print(f"결과 파일: {(OUTPUT_DIR / filename).resolve()}")


# =========================================================
# 12. 전체 문서 변환
# =========================================================

def convert_file(input_file):

    input_file = Path(input_file).resolve()
    extension = input_file.suffix.lower()
    if extension not in SUPPORTED_EXTENSIONS:
        raise ValueError("지원하지 않는 파일 형식입니다.")
    if not input_file.is_file():
        raise FileNotFoundError(f"입력 파일이 없습니다: {input_file}")

    INPUT_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)

    process = None
    document = None
    profile_dir = None
    log_file = None

    manifest = {
        "source_file": input_file.name,
        "engine": "LibreOffice Calc UNO",
        "sheets": [],
    }

    try:

        process, desktop, profile_dir, log_file = (
            start_libreoffice()
        )

        print(f"\n[3/4] {extension} 문서 열기...")

        document = open_document(
            desktop,
            input_file
        )

        if extension == ".docx":
            if not document.supportsService("com.sun.star.text.TextDocument"):
                raise ValueError("DOCX 파일을 Writer 문서로 열지 못했습니다.")
            print("\n[4/4] 문단 및 표를 Markdown으로 변환...")
            write_writer_output(document, input_file)
            return

        if not document.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
            raise ValueError("Excel 파일을 Calc 문서로 열지 못했습니다.")

        sheet_names = document.Sheets.getElementNames()

        if not sheet_names:
            raise RuntimeError("엑셀 파일에 시트가 없습니다.")

        index_lines = [
            f"# {input_file.stem}",
            "",
            "## 시트 목록",
            ""
        ]

        print(f"\n전체 시트: {len(sheet_names)}개")

        with tqdm(
            total=len(sheet_names),
            desc="시트 변환",
            unit="시트",
            dynamic_ncols=True
        ) as pbar:

            for index, sheet_name in enumerate(
                sheet_names,
                start=1
            ):

                pbar.set_postfix_str(
                    f"처리 중: {sheet_name}"
                )

                sheet = document.Sheets.getByName(
                    sheet_name
                )

                extracted = extract_sheet(sheet)

                blocks = find_blocks(
                    extracted["grid"]
                )

                markdown = blocks_to_markdown(
                    sheet_name,
                    blocks
                )

                filename = unique_filename(
                    index,
                    sheet_name
                )

                output_path = OUTPUT_DIR / filename

                output_path.write_text(
                    markdown,
                    encoding="utf-8"
                )

                index_lines.append(
                    f"- [{sheet_name}]({filename})"
                )

                manifest["sheets"].append({
                    "index": index,
                    "name": sheet_name,
                    "file": filename,
                    "used_area": {
                        "start_row": extracted["start_row"],
                        "end_row": extracted["end_row"],
                        "start_col": extracted["start_col"],
                        "end_col": extracted["end_col"],
                    },
                    "blocks": blocks,
                })

                pbar.update(1)

        print("\n[4/4] 통합 결과 파일 생성...")

        (OUTPUT_DIR / "workbook.md").write_text(
            "\n".join(index_lines) + "\n",
            encoding="utf-8"
        )

        (OUTPUT_DIR / "manifest.json").write_text(
            json.dumps(
                manifest,
                ensure_ascii=False,
                indent=2
            ),
            encoding="utf-8"
        )

        print("\n변환 완료!")
        print(f"변환 시트: {len(sheet_names)}개")
        print(f"결과 폴더: {OUTPUT_DIR.resolve()}")

    finally:

        if document is not None:

            try:
                document.close(True)
            except Exception:
                pass

        if process is not None:

            try:

                process.terminate()
                process.wait(timeout=10)

            except subprocess.TimeoutExpired:

                process.kill()
                process.wait()

        if log_file is not None:
            log_file.close()

        if profile_dir is not None:

            shutil.rmtree(
                profile_dir,
                ignore_errors=True
            )


# =========================================================
# 13. 실행 진입점
# =========================================================

def main():

    parser = argparse.ArgumentParser(
        description=(
            "LibreOffice를 이용한 "
            "Excel/DOCX 텍스트·표 Markdown 변환기 (이미지 제외)"
        )
    )

    parser.add_argument(
        "file",
        help="변환할 파일 경로 (.xlsx, .xls, .xlsm, .ods, .docx)"
    )

    args = parser.parse_args()

    input_file = Path(args.file).resolve()

    if not input_file.exists():

        raise FileNotFoundError(
            f"입력 파일이 없습니다: {input_file}"
        )

    if input_file.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise ValueError("지원하지 않는 파일 형식입니다.")

    convert_file(input_file)


if __name__ == "__main__":
    main()
