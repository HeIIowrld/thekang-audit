# -*- coding: utf-8 -*-
"""Small FastAPI UI for uploading source materials and downloading reviewer packages."""

from __future__ import annotations

import html
import json
import os
import re
import shutil
import uuid
import zipfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles


APP_ROOT = Path(__file__).resolve().parent
PROJECT_ROOT = APP_ROOT.parent
RUNS_ROOT = Path(os.getenv("DISCLOSURE_WEB_RUNS", PROJECT_ROOT / "web_runs")).resolve()
MAX_UPLOAD_BYTES = int(os.getenv("DISCLOSURE_WEB_MAX_UPLOAD_MB", "512")) * 1024 * 1024
APP_TITLE = "Disclosure Review Distributor"

app = FastAPI(title=APP_TITLE)
app.mount("/static", StaticFiles(directory=APP_ROOT / "static"), name="static")


@dataclass(frozen=True)
class JobPaths:
    root: Path
    uploads: Path
    source_files: Path
    packages: Path
    manifest: Path


def job_paths(job_id: str) -> JobPaths:
    if not re.fullmatch(r"[0-9a-f]{12}", job_id):
        raise HTTPException(status_code=404, detail="Unknown job")
    root = (RUNS_ROOT / job_id).resolve()
    if root != RUNS_ROOT and not root.is_relative_to(RUNS_ROOT):
        raise HTTPException(status_code=400, detail="Invalid job path")
    return JobPaths(
        root=root,
        uploads=root / "uploads",
        source_files=root / "source_files",
        packages=root / "packages",
        manifest=root / "job.json",
    )


def ensure_runs_root() -> None:
    RUNS_ROOT.mkdir(parents=True, exist_ok=True)


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def e(value: Any) -> str:
    return html.escape(str(value), quote=True)


def u(value: Any) -> str:
    return quote(str(value), safe="")


def page(title: str, body: str) -> HTMLResponse:
    return HTMLResponse(
        f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{e(title)} - {APP_TITLE}</title>
  <link rel="stylesheet" href="/static/styles.css">
</head>
<body>
  <header class="topbar">
    <a class="brand" href="/">Disclosure Review Distributor</a>
    <nav><a href="/">작업 목록</a></nav>
  </header>
  <main>{body}</main>
</body>
</html>"""
    )


def parse_reviewers(raw: str) -> list[str]:
    reviewers: list[str] = []
    for token in re.split(r"[\n,]+", raw):
        reviewer = token.strip()
        if reviewer and reviewer not in reviewers:
            reviewers.append(reviewer)
    return reviewers


def safe_name(name: str, fallback: str = "upload") -> str:
    name = Path(name or fallback).name
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip(" .")
    return name or fallback


def unique_child(parent: Path, filename: str) -> Path:
    filename = safe_name(filename)
    candidate = parent / filename
    if not candidate.exists():
        return candidate
    stem = candidate.stem
    suffix = candidate.suffix
    for index in range(2, 10000):
        candidate = parent / f"{stem}_{index}{suffix}"
        if not candidate.exists():
            return candidate
    raise RuntimeError(f"Could not create unique filename for {filename}")


def safe_extract_zip(zip_path: Path, target_dir: Path) -> list[Path]:
    extracted: list[Path] = []
    target_dir.mkdir(parents=True, exist_ok=True)
    target_resolved = target_dir.resolve()
    with zipfile.ZipFile(zip_path) as archive:
        for info in archive.infolist():
            if info.is_dir():
                continue
            parts = [safe_name(part, "item") for part in Path(info.filename).parts if part not in {"", ".", ".."}]
            if not parts:
                continue
            destination = (target_dir.joinpath(*parts)).resolve()
            if not destination.is_relative_to(target_resolved):
                raise HTTPException(status_code=400, detail=f"Unsafe zip entry: {info.filename}")
            destination.parent.mkdir(parents=True, exist_ok=True)
            with archive.open(info) as source, destination.open("wb") as target:
                shutil.copyfileobj(source, target)
            extracted.append(destination)
    return extracted


async def save_upload(upload: UploadFile, destination: Path) -> int:
    total = 0
    destination.parent.mkdir(parents=True, exist_ok=True)
    with destination.open("wb") as target:
        while True:
            chunk = await upload.read(1024 * 1024)
            if not chunk:
                break
            total += len(chunk)
            if total > MAX_UPLOAD_BYTES:
                raise HTTPException(status_code=413, detail=f"{upload.filename} is larger than the upload limit")
            target.write(chunk)
    await upload.close()
    return total


def relative_file_rows(base: Path) -> list[Path]:
    files = [path for path in base.rglob("*") if path.is_file()]
    return sorted(files, key=lambda path: path.relative_to(base).as_posix().lower())


def reviewer_match(path: Path, base: Path, reviewers: list[str]) -> str | None:
    relative = path.relative_to(base).as_posix().lower()
    tokens = re.findall(r"[0-9a-z가-힣]+", relative)
    for reviewer in reviewers:
        needle = reviewer.lower()
        if needle in tokens:
            return reviewer
        if len(needle) >= 3 and needle in relative:
            return reviewer
    return None


def assign_files(files: list[Path], base: Path, reviewers: list[str], mode: str) -> dict[str, list[Path]]:
    assignments: dict[str, list[Path]] = {reviewer: [] for reviewer in reviewers}
    fallback_index = 0
    for path in files:
        reviewer = reviewer_match(path, base, reviewers) if mode == "name_match" else None
        if reviewer is None:
            reviewer = reviewers[fallback_index % len(reviewers)]
            fallback_index += 1
        assignments[reviewer].append(path)
    return assignments


def write_package(job: JobPaths, reviewer: str, files: list[Path]) -> Path:
    zip_path = job.packages / f"{safe_name(reviewer, 'reviewer')}_materials.zip"
    zip_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        readme = (
            f"Reviewer: {reviewer}\n"
            f"Generated: {now_text()}\n"
            f"Files: {len(files)}\n\n"
            "이 압축파일은 웹 UI에서 자동 배분한 자료 묶음입니다.\n"
        )
        archive.writestr("README_배분안내.txt", readme)
        for source in files:
            archive.write(source, source.relative_to(job.source_files).as_posix())
    return zip_path


def load_job(job_id: str) -> dict[str, Any]:
    job = job_paths(job_id)
    if not job.manifest.exists():
        raise HTTPException(status_code=404, detail="Unknown job")
    return json.loads(job.manifest.read_text(encoding="utf-8"))


def save_manifest(job: JobPaths, payload: dict[str, Any]) -> None:
    job.manifest.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def list_jobs() -> list[dict[str, Any]]:
    ensure_runs_root()
    jobs: list[dict[str, Any]] = []
    for manifest in RUNS_ROOT.glob("*/job.json"):
        try:
            payload = json.loads(manifest.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        jobs.append(payload)
    return sorted(jobs, key=lambda item: item.get("created_at", ""), reverse=True)


def format_bytes(value: int) -> str:
    size = float(value)
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:.1f} {unit}" if unit != "B" else f"{int(size)} B"
        size /= 1024
    return f"{value} B"


def upload_form(error: str = "") -> str:
    error_html = f'<div class="alert">{e(error)}</div>' if error else ""
    return f"""
<section class="panel">
  <div class="panel-heading">
    <h1>자료 배분</h1>
    <p>원자료를 올리고 담당자별 ZIP 패키지를 생성합니다.</p>
  </div>
  {error_html}
  <form class="upload-form" action="/jobs" method="post" enctype="multipart/form-data">
    <label>
      작업명
      <input name="title" type="text" placeholder="예: 1차 검토 자료 배분" maxlength="80">
    </label>
    <label>
      담당자
      <textarea name="reviewers" rows="4" placeholder="A&#10;B&#10;C" required></textarea>
    </label>
    <div class="form-grid">
      <label>
        배분 방식
        <select name="assignment_mode">
          <option value="round_robin">라운드로빈</option>
          <option value="name_match">파일/폴더명 담당자 매칭 후 나머지 라운드로빈</option>
        </select>
      </label>
      <label class="checkbox">
        <input type="checkbox" name="extract_archives" value="yes" checked>
        ZIP 파일은 풀어서 배분
      </label>
    </div>
    <label>
      업로드 자료
      <input name="files" type="file" multiple required>
    </label>
    <button type="submit">패키지 생성</button>
  </form>
</section>
"""


@app.get("/", response_class=HTMLResponse)
def index() -> HTMLResponse:
    jobs = list_jobs()
    rows = "\n".join(
        f"""<tr>
  <td><a href="/jobs/{e(job['id'])}">{e(job.get('title') or job['id'])}</a></td>
  <td>{e(job.get('created_at', ''))}</td>
  <td>{len(job.get('reviewers', []))}</td>
  <td>{job.get('file_count', 0)}</td>
  <td>{job.get('total_size_label', '-')}</td>
</tr>"""
        for job in jobs
    ) or '<tr><td colspan="5" class="empty">아직 생성된 작업이 없습니다.</td></tr>'
    body = upload_form() + f"""
<section class="panel">
  <div class="panel-heading">
    <h2>작업 목록</h2>
    <p>생성된 작업을 열어 담당자별 패키지를 내려받습니다.</p>
  </div>
  <div class="table-wrap">
    <table>
      <thead><tr><th>작업명</th><th>생성일</th><th>담당자</th><th>파일</th><th>용량</th></tr></thead>
      <tbody>{rows}</tbody>
    </table>
  </div>
</section>
"""
    return page("작업 목록", body)


@app.post("/jobs")
async def create_job(
    title: str = Form(""),
    reviewers: str = Form(...),
    assignment_mode: str = Form("round_robin"),
    extract_archives: str | None = Form(None),
    files: list[UploadFile] = File(...),
) -> Any:
    reviewer_list = parse_reviewers(reviewers)
    if not reviewer_list:
        return page("입력 오류", upload_form("담당자를 한 명 이상 입력하세요."))
    if assignment_mode not in {"round_robin", "name_match"}:
        return page("입력 오류", upload_form("지원하지 않는 배분 방식입니다."))

    ensure_runs_root()
    job_id = uuid.uuid4().hex[:12]
    job = job_paths(job_id)
    job.uploads.mkdir(parents=True, exist_ok=True)
    job.source_files.mkdir(parents=True, exist_ok=True)
    job.packages.mkdir(parents=True, exist_ok=True)

    uploaded: list[dict[str, Any]] = []
    total_bytes = 0
    for upload in files:
        if not upload.filename:
            continue
        upload_path = unique_child(job.uploads, upload.filename)
        size = await save_upload(upload, upload_path)
        total_bytes += size
        uploaded.append({"name": upload_path.name, "size": size})
        if extract_archives == "yes" and upload_path.suffix.lower() == ".zip":
            safe_extract_zip(upload_path, job.source_files / upload_path.stem)
        else:
            target = unique_child(job.source_files, upload_path.name)
            shutil.copy2(upload_path, target)

    source_files = relative_file_rows(job.source_files)
    if not source_files:
        shutil.rmtree(job.root, ignore_errors=True)
        return page("입력 오류", upload_form("배분할 파일을 찾지 못했습니다."))

    assignments = assign_files(source_files, job.source_files, reviewer_list, assignment_mode)
    packages: list[dict[str, Any]] = []
    for reviewer, assigned_files in assignments.items():
        package_path = write_package(job, reviewer, assigned_files)
        packages.append(
            {
                "reviewer": reviewer,
                "file_count": len(assigned_files),
                "zip_name": package_path.name,
                "zip_size": package_path.stat().st_size,
            }
        )

    manifest = {
        "id": job_id,
        "title": title.strip() or f"배분 작업 {job_id}",
        "created_at": now_text(),
        "reviewers": reviewer_list,
        "assignment_mode": assignment_mode,
        "extract_archives": extract_archives == "yes",
        "uploads": uploaded,
        "file_count": len(source_files),
        "total_size": total_bytes,
        "total_size_label": format_bytes(total_bytes),
        "packages": packages,
        "assignments": {
            reviewer: [path.relative_to(job.source_files).as_posix() for path in assigned_files]
            for reviewer, assigned_files in assignments.items()
        },
    }
    save_manifest(job, manifest)
    return RedirectResponse(f"/jobs/{job_id}", status_code=303)


@app.get("/jobs/{job_id}", response_class=HTMLResponse)
def job_detail(job_id: str) -> HTMLResponse:
    data = load_job(job_id)
    package_rows = "\n".join(
        f"""<tr>
  <td>{e(package['reviewer'])}</td>
  <td>{package['file_count']}</td>
  <td>{format_bytes(int(package.get('zip_size', 0)))}</td>
  <td><a class="button small" href="/jobs/{u(job_id)}/download?reviewer={u(package['reviewer'])}">다운로드</a></td>
</tr>"""
        for package in data.get("packages", [])
    )
    more_text = "\n..."
    assignment_rows = "\n".join(
        f"""<tr>
  <td>{e(reviewer)}</td>
  <td>{len(files)}</td>
  <td class="path-list">{e(chr(10).join(files[:20]))}{e(more_text) if len(files) > 20 else ''}</td>
</tr>"""
        for reviewer, files in data.get("assignments", {}).items()
    )
    body = f"""
<section class="panel">
  <div class="panel-heading with-actions">
    <div>
      <h1>{e(data.get('title', job_id))}</h1>
      <p>생성일 {e(data.get('created_at', ''))} · 업로드 {e(data.get('total_size_label', '-'))} · 파일 {data.get('file_count', 0)}개</p>
    </div>
    <form action="/jobs/{e(job_id)}/delete" method="post">
      <button class="secondary danger" type="submit">작업 삭제</button>
    </form>
  </div>
  <div class="table-wrap">
    <table>
      <thead><tr><th>담당자</th><th>파일 수</th><th>ZIP 크기</th><th>다운로드</th></tr></thead>
      <tbody>{package_rows}</tbody>
    </table>
  </div>
</section>
<section class="panel">
  <div class="panel-heading">
    <h2>배분 내역</h2>
    <p>각 담당자 ZIP에 들어간 원본 파일 경로입니다.</p>
  </div>
  <div class="table-wrap">
    <table>
      <thead><tr><th>담당자</th><th>파일 수</th><th>파일</th></tr></thead>
      <tbody>{assignment_rows}</tbody>
    </table>
  </div>
</section>
"""
    return page(str(data.get("title", job_id)), body)


@app.get("/jobs/{job_id}/download")
def download_package(job_id: str, reviewer: str) -> FileResponse:
    data = load_job(job_id)
    package = next((item for item in data.get("packages", []) if item.get("reviewer") == reviewer), None)
    if not package:
        raise HTTPException(status_code=404, detail="Unknown reviewer package")
    job = job_paths(job_id)
    zip_path = (job.packages / package["zip_name"]).resolve()
    if not zip_path.is_relative_to(job.packages.resolve()) or not zip_path.exists():
        raise HTTPException(status_code=404, detail="Package not found")
    return FileResponse(zip_path, filename=zip_path.name, media_type="application/zip")


@app.post("/jobs/{job_id}/delete")
def delete_job(job_id: str) -> RedirectResponse:
    job = job_paths(job_id)
    if job.root.exists():
        shutil.rmtree(job.root)
    return RedirectResponse("/", status_code=303)
