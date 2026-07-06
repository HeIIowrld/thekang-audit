# Disclosure Review Queue Toolkit

공시자료, 제출 엑셀, 근거자료를 대조해 사람이 확인할 검토 큐를 만드는 로컬 자동화 도구입니다. 이 프로젝트는 최종 판정을 자동 확정하지 않고, 검토자가 원문 링크와 후보 사유를 확인할 수 있도록 후보를 줄이는 데 초점을 둡니다.

## 공개 저장소 원칙

이 저장소에는 코드와 문서만 올립니다. 실제 기관 제출자료, 증빙, 원문 PDF, 산출 엑셀, 실행 로그, 개인별 배정 폴더는 `.gitignore`로 제외합니다.

공개 전에는 아래 명령으로 추적 대상에 데이터 파일이 섞이지 않았는지 확인하세요.

```powershell
git status --short
git check-ignore -v "30_검토산출물/01_검토시작/01_검토큐.xlsx"
```

## 입력 구조

기본 파이프라인은 로컬 작업공간에 다음 구조가 있다고 가정합니다.

```text
project-root/
  20_기준자료/
    02_배정완료_원본/
  30_검토산출물/
    90_실행로그/
      90_재실행캐시/
  90_tools/
```

자세한 입력 배치와 파일명 필터는 [docs/INPUT_STRUCTURE.md](docs/INPUT_STRUCTURE.md)를 참고하세요.

## 설치

Python 3.11 이상을 권장합니다.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
```

## 실행

최종 검토 큐와 전달 패키지를 생성합니다.

```powershell
python .\90_tools\run_review_pipeline.py --root .
```

이미 생성된 자동검토 캐시를 재사용해 경량 패키지만 다시 만들 수 있습니다.

```powershell
python .\90_tools\run_review_pipeline.py --root . --use-existing-auto
```

산출물 구조만 점검하려면 아래 명령을 사용합니다.

```powershell
python .\90_tools\run_review_pipeline.py --root . --audit-only
```

## 웹 UI

자료 업로드와 담당자별 ZIP 다운로드가 필요한 경우 웹 UI를 실행합니다.

```powershell
python -m uvicorn web_app.app:app --reload --host 127.0.0.1 --port 8000
```

브라우저에서 `http://127.0.0.1:8000`을 열고 담당자 목록과 자료를 업로드하면 담당자별 다운로드 패키지가 생성됩니다. 해당 포트가 이미 사용 중이면 `--port 8001`처럼 다른 포트를 지정합니다. 업로드 자료와 생성 ZIP은 `web_runs/`에 저장되며 Git에는 포함되지 않습니다.

자세한 사용법은 [docs/WEB_UI.md](docs/WEB_UI.md)를 참고하세요.

## 주요 산출물

```text
30_검토산출물/01_검토시작/01_검토큐.xlsx
30_검토산출물/90_실행로그/01_파이프라인_점검.md
40_전달패키지/01_검토시작/01_검토큐.xlsx
```

`01_검토큐.xlsx`는 검토자가 상태, 판정, 재확인 여부, 검토메모를 입력하는 작업 파일입니다. 링크 열에서 제출 엑셀, 공시 원문, 근거자료를 열어 확인한 뒤 필요한 행만 결과표 초안에 반영합니다.

## 일반화 옵션

제출 엑셀 파일명이 특정 업무명과 맞지 않으면 파일명 필터를 바꿀 수 있습니다.

```powershell
python .\90_tools\run_review_pipeline.py --root . --submitted-name-keyword "review,disclosure,checklist"
```

파일명 대신 시트 구조만 보고 제출 엑셀을 찾으려면 아래 옵션을 사용합니다.

```powershell
python .\90_tools\run_review_pipeline.py --root . --disable-submitted-name-filter
```

담당자 기본 목록은 환경변수로 지정할 수 있습니다.

```powershell
$env:DISCLOSURE_REVIEWERS = "A,B,C,D"
```

입력 자료에 이미 다른 담당자명이 들어 있으면 파이프라인이 해당 담당자도 자동으로 포함합니다.

## 주의

이 도구는 공개 가능한 샘플 데이터 없이도 저장소를 배포할 수 있도록 구성되어 있습니다. 실제 데이터로 실행한 뒤에는 `git status --short`에서 엑셀, PDF, HWP, CSV, ZIP 등이 보이지 않는지 확인하세요.
