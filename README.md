# Excel·Word 문서 Markdown 변환기

LibreOffice의 UNO 인터페이스를 이용해 스프레드시트와 Word 문서를 Markdown으로 변환합니다. 문서의 텍스트와 표를 추출하며, 원본의 시각적 서식을 그대로 재현하는 도구는 아닙니다.

## 지원 파일

- 스프레드시트: `.xlsx`, `.xls`, `.xlsm`, `.ods`
- Word 문서: `.docx`

확장자는 대소문자를 구분하지 않습니다.

## 프로그램 설치

### 1. Python 설치

1. [Python 공식 다운로드 페이지](https://www.python.org/downloads/windows/)에서 Python 3.13의 64-bit 설치 프로그램을 내려받아 실행합니다. `requirements.txt`에는 Python 3.13.17이 기준 버전으로 적혀 있습니다.
2. 설치 첫 화면에서 **Add python.exe to PATH**를 선택한 뒤 설치합니다.
3. PowerShell을 새로 열어 설치를 확인합니다.

```powershell
python --version
```

Python 3.13이 표시되지 않으면 Windows Python Launcher로 확인할 수 있습니다.

```powershell
py -3.13 --version
```

### 2. LibreOffice 설치

[LibreOffice 공식 다운로드 페이지](https://www.libreoffice.org/download/download-libreoffice/)에서 Windows용 LibreOffice를 내려받아 기본 옵션으로 설치합니다. 프로그램은 기본적으로 다음 경로에 설치되어 있어야 합니다.

```text
C:\Program Files\LibreOffice\program\soffice.exe
```

스크립트는 이 경로를 우선 확인하고, 없으면 `C:\Program Files (x86)\LibreOffice\program\soffice.exe`를 확인합니다. 다른 위치에 설치했다면 [extract_to_markdown.py](./extract_to_markdown.py)의 `LO_PROGRAM` 경로를 실제 설치 경로로 바꿔야 합니다.

### 3. 프로젝트 가상환경 및 패키지 설치

PowerShell에서 프로젝트 폴더로 이동합니다. 예를 들어 프로젝트가 `C:\개인프로젝트`에 있다면:

```powershell
Set-Location C:\개인프로젝트
py -3.13 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

`py` 명령을 사용할 수 없다면 Python 설치 후 `python -m venv .venv`를 사용하고, 이후 명령에서도 `.\.venv\Scripts\python.exe`를 실행하면 됩니다. 가상환경을 활성화하지 않아도 위 명령으로 설치와 실행이 가능합니다.

### 4. LibreOffice UNO 연결 확인

UNO는 LibreOffice와 함께 제공되는 인터페이스이며, 별도의 `pip install uno` 패키지를 설치하는 방식이 아닙니다. 가상환경에서 모듈을 가져올 수 있는지 확인합니다.

```powershell
.\.venv\Scripts\python.exe -c "import extract_to_markdown; print('UNO 연결 준비 완료')"
```

오류가 나면 LibreOffice 설치 경로와 Python 버전을 확인하세요. LibreOffice의 UNO 모듈과 현재 Python 환경이 호환되지 않으면 모듈을 불러오지 못할 수 있습니다.

### 5. 첫 변환 실행

```powershell
.\.venv\Scripts\python.exe .\extract_to_markdown.py ".\input\문서.xlsx"
```

변환 결과는 프로젝트의 `output` 폴더에 생성됩니다. 실행 중 LibreOffice 서버가 시작되고, 작업이 끝나면 자동 종료됩니다.

## 사용 방법

변환할 파일 경로를 인자로 지정합니다. 파일 경로에 공백이 있으면 따옴표로 감쌉니다.

```powershell
.\.venv\Scripts\python.exe .\extract_to_markdown.py ".\input\문서.xlsx"
.\.venv\Scripts\python.exe .\extract_to_markdown.py ".\input\보고서.docx"
```

입력 파일은 `input` 폴더 안에 둘 필요는 없습니다. 지정한 경로에서 읽으며, 결과는 스크립트가 있는 폴더의 `output`에 저장합니다. 같은 이름의 결과 파일이 있으면 덮어씁니다.

## 결과물

### 스프레드시트

- `sheet_01_시트명.md` 등: 시트별 Markdown
- `workbook.md`: 시트별 파일 링크가 있는 목차
- `manifest.json`: 원본 파일명, 시트 정보, 추출 블록 및 사용 영역

### Word 문서

- `문서명.md`: 문단과 표를 문서 순서대로 변환한 결과
- `manifest.json`: 원본 파일명과 추출 블록 정보

`manifest.json`의 구조는 스프레드시트와 Word 문서에서 서로 다릅니다. 두 형식 모두 같은 `output` 폴더를 사용하므로, 각 실행은 해당 형식의 manifest를 새로 씁니다. 변환 전에 폴더의 기존 파일을 자동 삭제하지는 않습니다.

## 테스트

기본 단위 테스트:

```powershell
.\.venv\Scripts\python.exe -m unittest -v
```

설치된 LibreOffice를 실행해 실제 변환까지 확인하려면 다음처럼 실행합니다. 이 테스트는 임시 파일을 만들고 LibreOffice 프로세스를 띄우므로 LibreOffice가 정상 설치되어 있어야 합니다.

```powershell
$env:RUN_LIBREOFFICE_TESTS = "1"
.\.venv\Scripts\python.exe -m unittest -v
Remove-Item Env:RUN_LIBREOFFICE_TESTS
```

## 현재 한계

- **환경 의존성:** LibreOffice 설치 경로가 코드에 고정되어 있고, Python에서 사용할 수 있는 UNO 바인딩이 필요합니다. Python/LibreOffice 조합에 따라 추가 환경 설정이 필요할 수 있습니다.
- **포트 충돌:** UNO 통신에 로컬 포트 `2002`를 고정 사용합니다. 같은 포트를 다른 LibreOffice 프로세스가 사용 중이면 연결에 실패할 수 있습니다.
- **스프레드시트 구조 판별:** 두 개 이상의 값이 있는 연속 행을 표 후보로 보는 규칙 기반 방식입니다. 첫 행을 헤더로 취급하므로 표가 아닌 데이터나 헤더 없는 표는 잘못 분류될 수 있습니다.
- **레이아웃·서식 손실:** 셀의 색, 글꼴, 너비, 병합 형태, 테두리, 수식 자체와 같은 시각·구조 정보는 보존하지 않습니다. 셀 값과 문서 텍스트·표를 Markdown에 맞게 단순화합니다.
- **개체 미지원:** 이미지 및 기타 삽입 개체는 추출하지 않습니다. 차트, 도형, 주석, 하이퍼링크 등도 원형대로 보존되지 않습니다.
- **대용량 시트:** 시트의 사용 영역을 한 번에 메모리에 읽고 처리합니다. 서식만 적용된 셀 때문에 사용 영역이 커진 파일은 처리 속도와 메모리 사용량이 증가할 수 있습니다.
- **출력 관리:** 결과 경로를 지정할 수 없고 모든 실행이 같은 `output` 폴더를 씁니다. 이전 실행에서 남은 파일을 정리하지 않으며, 변환이 중간에 실패하면 일부 결과가 남을 수 있습니다.

## 다음 개선 단계

1. **실행 설정 분리:** 입력·출력 폴더, LibreOffice 경로, UNO 포트를 명령행 옵션이나 설정으로 지정하고 포트 사용 가능 여부를 확인합니다.
2. **출력 안전성 강화:** 원본별 출력 폴더를 지원하고, 변환 성공 후 결과를 확정해 부분 생성 및 파일 덮어쓰기 혼선을 줄입니다.
3. **추출 정확도 개선:** 표 판별 규칙을 보완하고 헤더 유무·표 여부를 사용자가 지정할 수 있게 합니다. 수식 결과 및 셀 좌표 등 필요한 정보를 선택적으로 기록합니다.
4. **지원 범위와 테스트 확장:** 대용량·빈 시트·복잡한 병합 셀·특수 문자와 형식별 오류 사례를 테스트 픽스처로 추가하고, 환경 문제를 구분하는 안내를 개선합니다.
5. **개체 처리 검토:** 사용 수요를 확인한 뒤 이미지·차트 등은 Markdown 링크나 별도 파일로 내보내는 방식을 검토합니다.
