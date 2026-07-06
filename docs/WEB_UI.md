# Web UI

웹 UI는 자료 업로드와 담당자별 다운로드 패키지 생성을 처리하는 작은 로컬 앱입니다. 기존 검토 파이프라인을 대체하지 않고, 원자료를 담당자에게 나눠 전달하는 작업을 단순화합니다.

## 실행

```powershell
python -m pip install -r requirements.txt
python -m uvicorn web_app.app:app --reload --host 127.0.0.1 --port 8000
```

브라우저에서 아래 주소를 엽니다.

```text
http://127.0.0.1:8000
```

포트 8000이 이미 사용 중이면 `--port 8001`처럼 다른 포트를 지정합니다.

## 사용 흐름

1. 작업명을 입력합니다.
2. 담당자명을 줄바꿈 또는 쉼표로 입력합니다.
3. 배분 방식을 선택합니다.
4. 파일 또는 ZIP을 업로드합니다.
5. 생성된 작업 상세 화면에서 담당자별 ZIP을 내려받습니다.

## 배분 방식

`라운드로빈`은 정렬된 파일 목록을 담당자 순서대로 배분합니다.

`파일/폴더명 담당자 매칭`은 파일 경로나 파일명에 담당자명이 포함된 경우 해당 담당자에게 우선 배분하고, 매칭되지 않은 파일은 라운드로빈으로 배분합니다.

## 저장 위치

업로드 파일과 생성된 ZIP은 기본적으로 아래에 저장됩니다.

```text
web_runs/
```

저장 위치를 바꾸려면 환경변수를 지정합니다.

```powershell
$env:DISCLOSURE_WEB_RUNS = "D:\review-web-runs"
```

업로드 크기 제한은 기본 512MB입니다.

```powershell
$env:DISCLOSURE_WEB_MAX_UPLOAD_MB = "2048"
```

`web_runs/`는 `.gitignore`에 포함되어 있어 GitHub에 올라가지 않습니다.
