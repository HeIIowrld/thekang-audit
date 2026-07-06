# Input Structure

이 저장소는 실데이터를 포함하지 않습니다. 실행할 때는 각자 로컬 작업공간에 원천 자료와 산출물 폴더를 준비해야 합니다.

## 기본 폴더

```text
project-root/
  20_기준자료/
    02_배정완료_원본/
      <institution-or-reviewer-folders>/
  30_검토산출물/
    90_실행로그/
      90_재실행캐시/
        01_자동검토/
        02_기준입력/
  90_tools/
```

## 공시 원천 엑셀

`disclosure_auto_review.py`는 기본적으로 루트 하위 폴더 중 `*.xlsx`가 충분히 있고, `5-1`, `6-2`, `10`, `13-1`, `13-5`로 시작하는 파일이 있는 폴더를 공시 원천 폴더로 추정합니다.

명시적으로 지정하려면 개별 스크립트를 아래처럼 실행합니다.

```powershell
python .\90_tools\disclosure_auto_review.py --root . --alio-dir .\path\to\public-disclosure-exports
```

## 제출 엑셀

기본 제출자료 루트는 `20_*/02_*` 폴더입니다. 파일명에는 기본적으로 아래 키워드 중 하나가 포함되어야 하고, 대상 시트가 4개 이상 있어야 합니다.

```text
review, disclosure, checklist, 노무, 통합공시, 점검표
```

다른 파일명 규칙을 쓰는 경우:

```powershell
python .\90_tools\run_review_pipeline.py --root . --submitted-name-keyword "my-review,my-checklist"
```

파일명 필터 없이 시트 구조만 사용하려는 경우:

```powershell
python .\90_tools\run_review_pipeline.py --root . --disable-submitted-name-filter
```

## 담당자 목록

담당자별 큐 파일명과 안내문에 쓰는 기본 담당자 목록은 환경변수로 지정합니다.

```powershell
$env:DISCLOSURE_REVIEWERS = "ReviewerA,ReviewerB,ReviewerC,ReviewerD"
```

지정하지 않으면 `A,B,C,D`를 사용합니다.

관리 큐나 자동검토 중간 산출물에 다른 담당자명이 이미 들어 있으면 해당 담당자도 자동으로 포함됩니다.

## 공개 전 확인

아래 유형의 파일은 GitHub에 올리지 않습니다.

```text
*.xlsx, *.pdf, *.hwp, *.hwpx, *.docx, *.csv, *.zip
20_기준자료/, 30_검토산출물/, 40_전달패키지/
```

공개 전 확인:

```powershell
git status --short
```
