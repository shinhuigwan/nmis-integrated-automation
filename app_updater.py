"""Verified GitHub Releases updates; user data is never part of the payload."""
import hashlib
import json
import os
import re
import shutil
import stat
import subprocess
import tempfile
import urllib.request
import zipfile
from pathlib import Path, PurePosixPath

REPOSITORY = "shinhuigwan/nmis-integrated-automation"
ASSET_NAME = "NMIS-Automation-Windows-x64.zip"
ROOT_NAME = "NMIS-Automation"
PAYLOAD_NAMES = {"NMIS-Automation.exe", "_internal", "installers", "Setup.bat", "Setup.ps1",
                 "Start.vbs", "run_ui.vbs", "apply_update.ps1", "사용방법.txt",
                 "THIRD_PARTY_LICENSES.txt", "BUILD_INFO.json"}


def version_tuple(value):
    value = str(value).lstrip("v")
    if not re.fullmatch(r"\d+(?:\.\d+){2,3}", value):
        raise ValueError("업데이트 버전 형식이 올바르지 않습니다.")
    numbers = tuple(int(piece) for piece in value.split("."))
    return numbers + (0,) * (4 - len(numbers))


def asset_url(url, tag):
    expected = f"https://github.com/{REPOSITORY}/releases/download/{tag}/"
    if not url.startswith(expected) or "?" in url or "#" in url:
        raise ValueError("업데이트 다운로드 주소가 올바르지 않습니다.")
    return url


def request(url):
    return urllib.request.urlopen(urllib.request.Request(url, headers={
        "User-Agent": "NMIS-Automation-Updater", "Accept": "application/vnd.github+json" if "api.github.com" in url else "application/octet-stream"}), timeout=30)


def check_latest(current_version):
    with request(f"https://api.github.com/repos/{REPOSITORY}/releases/latest") as response:
        release = json.loads(response.read(2_000_000))
    tag = release["tag_name"]
    latest = tag.lstrip("v")
    if version_tuple(latest) <= version_tuple(current_version):
        return None
    assets = {item["name"]: item["browser_download_url"] for item in release.get("assets", [])}
    if ASSET_NAME not in assets or ASSET_NAME + ".sha256" not in assets:
        raise RuntimeError("새 버전의 업로드가 아직 끝나지 않았습니다. 잠시 후 다시 확인하세요.")
    return dict(version=latest, tag=tag, notes=str(release.get("body") or "")[:20000],
                url=asset_url(assets[ASSET_NAME], tag), checksum_url=asset_url(assets[ASSET_NAME + ".sha256"], tag))


def extract_payload(archive, destination):
    with zipfile.ZipFile(archive) as package:
        entries = package.infolist()
        if len(entries) > 30000 or sum(item.file_size for item in entries) > 2_000_000_000:
            raise ValueError("업데이트 파일 크기가 허용 범위를 초과했습니다.")
        for entry in entries:
            path = PurePosixPath(entry.filename)
            if ("\\" in entry.filename or path.is_absolute() or not path.parts or path.parts[0] != ROOT_NAME
                    or any(part in {".", ".."} or ":" in part or part != part.rstrip(" .") for part in path.parts)
                    or stat.S_ISLNK(entry.external_attr >> 16)):
                raise ValueError("업데이트 ZIP 경로를 확인하지 못했습니다.")
            if len(path.parts) > 1 and path.parts[1] not in PAYLOAD_NAMES:
                raise ValueError("업데이트에 사용자 데이터 또는 알 수 없는 파일이 포함되어 있습니다.")
        package.extractall(destination)
    root = Path(destination) / ROOT_NAME
    if not (root / "NMIS-Automation.exe").is_file() or not (root / "_internal").is_dir():
        raise ValueError("업데이트 실행 파일 또는 라이브러리 폴더가 없습니다.")
    return root


def stage_update(release, progress=lambda message: None):
    stage = Path(tempfile.mkdtemp(prefix="nmis-update-"))
    try:
        with request(release["checksum_url"]) as response:
            expected = response.read(4096).decode("ascii").split()[0].lower()
        if not re.fullmatch(r"[0-9a-f]{64}", expected):
            raise ValueError("업데이트 검증 정보를 읽지 못했습니다.")
        archive = stage / ASSET_NAME
        digest = hashlib.sha256()
        with request(release["url"]) as response, archive.open("wb") as output:
            total = int(response.headers.get("Content-Length") or 0)
            downloaded = 0
            while chunk := response.read(1024 * 1024):
                downloaded += len(chunk)
                if downloaded > 1_000_000_000:
                    raise ValueError("업데이트 다운로드 크기가 허용 범위를 초과했습니다.")
                output.write(chunk)
                digest.update(chunk)
                progress(f"다운로드 중: {downloaded // 1048576}MB" + (f" / {total // 1048576}MB" if total else ""))
        if digest.hexdigest() != expected:
            raise ValueError("업데이트 파일 검증에 실패했습니다. 현재 버전은 그대로 유지됩니다.")
        progress("파일 검증 완료. 실행 환경을 확인합니다...")
        root = extract_payload(archive, stage / "payload")
        report = stage / "probe.json"
        completed = subprocess.run([str(root / "NMIS-Automation.exe"), "--update-probe", str(report)],
                                   cwd=root, timeout=90, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        result = json.loads(report.read_text(encoding="utf-8")) if report.is_file() else {}
        if completed.returncode or not result.get("ok") or version_tuple(result.get("version", "0.0.0")) != version_tuple(release["version"]):
            raise RuntimeError("새 버전 실행 검증에 실패했습니다. 현재 버전은 그대로 유지됩니다.")
        return stage, root
    except Exception:
        shutil.rmtree(stage, ignore_errors=True)
        raise


def launch_update_helper(stage, payload, app_root):
    # Always use the current trusted helper; never an unverified downloaded script.
    root = Path(app_root).resolve()
    helper = root / "apply_update.ps1"
    if not helper.is_file():
        helper = root / "packaging" / "apply_update.ps1"
    if not helper.is_file():
        raise RuntimeError("업데이트 도우미 파일을 찾지 못했습니다.")
    staged_helper = Path(stage) / "apply_update.ps1"
    shutil.copy2(helper, staged_helper)
    subprocess.Popen([str(Path(os.environ["WINDIR"]) / "System32/WindowsPowerShell/v1.0/powershell.exe"),
                      "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(staged_helper),
                      "-AppRoot", str(root), "-PayloadRoot", str(Path(payload).resolve()),
                      "-ParentPid", str(os.getpid())], creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
