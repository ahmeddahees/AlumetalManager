import hashlib
import http.client
import json
import os
import re
import shutil
import subprocess
import tempfile
import uuid
from dataclasses import dataclass
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen

from github_config import GITHUB_OWNER, GITHUB_REPOSITORY


REQUEST_TIMEOUT = 15
MAX_EXECUTABLE_SIZE = 500 * 1024 * 1024
EXECUTABLE_NAME = "AlumetalManager.exe"
VERSION_PATTERN = re.compile(r"^v?(\d+)\.(\d+)\.(\d+)$")


class UpdateError(Exception):
    pass


class NoReleaseFound(Exception):
    pass


@dataclass(frozen=True)
class Release:
    version: str
    page_url: str
    executable_url: str
    checksum_url: str
    executable_size: int


def _version_key(value):
    match = VERSION_PATTERN.fullmatch(value)
    if not match:
        raise UpdateError(f"Unsupported version format: {value}")
    return tuple(int(part) for part in match.groups())


def _github_url(value):
    try:
        parsed = urlparse(value)
        hostname = parsed.hostname
    except ValueError as error:
        raise UpdateError("GitHub returned an invalid URL.") from error
    if parsed.scheme != "https" or hostname not in (
        "github.com",
        "api.github.com",
    ):
        raise UpdateError("GitHub returned an unexpected download URL.")
    return value


def get_latest_release(current_version):
    api_url = (
        f"https://api.github.com/repos/{GITHUB_OWNER}/"
        f"{GITHUB_REPOSITORY}/releases/latest"
    )
    request = Request(
        api_url,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": "AlumetalManager-Updater",
        },
    )
    try:
        with urlopen(request, timeout=REQUEST_TIMEOUT) as response:
            data = json.loads(response.read().decode("utf-8"))
    except HTTPError as error:
        if error.code == 404:
            raise NoReleaseFound from error
        raise UpdateError(f"GitHub returned HTTP {error.code}.") from error
    except (
        URLError,
        TimeoutError,
        http.client.HTTPException,
        json.JSONDecodeError,
        UnicodeDecodeError,
    ) as error:
        raise UpdateError(f"Could not check GitHub for updates: {error}") from error

    if not isinstance(data, dict):
        raise UpdateError("GitHub returned an invalid release response.")
    tag = data.get("tag_name")
    if not isinstance(tag, str):
        raise UpdateError("The latest release has no version tag.")
    latest_key = _version_key(tag)
    if latest_key <= _version_key(current_version):
        return None

    assets = data.get("assets")
    if not isinstance(assets, list):
        raise UpdateError("The latest release does not contain downloadable files.")
    asset_urls = {}
    executable_size = None
    for asset in assets:
        if not isinstance(asset, dict):
            continue
        name = asset.get("name")
        download_url = asset.get("browser_download_url")
        if name in (EXECUTABLE_NAME, "SHA256SUMS.txt") and isinstance(download_url, str):
            asset_urls[name] = _github_url(download_url)
            if name == EXECUTABLE_NAME:
                executable_size = asset.get("size")

    if executable_size is None or EXECUTABLE_NAME not in asset_urls:
        raise UpdateError(f"The latest release is missing {EXECUTABLE_NAME}.")
    if "SHA256SUMS.txt" not in asset_urls:
        raise UpdateError("The latest release is missing its checksum file.")
    if (
        not isinstance(executable_size, int)
        or executable_size <= 0
        or executable_size > MAX_EXECUTABLE_SIZE
    ):
        raise UpdateError("The latest release has an invalid executable size.")

    page_url = data.get("html_url")
    if not isinstance(page_url, str):
        raise UpdateError("The latest release has no web page URL.")
    return Release(
        version=tag.lstrip("v"),
        page_url=_github_url(page_url),
        executable_url=asset_urls[EXECUTABLE_NAME],
        checksum_url=asset_urls["SHA256SUMS.txt"],
        executable_size=executable_size,
    )


def _read_checksums(url):
    request = Request(url, headers={"User-Agent": "AlumetalManager-Updater"})
    with urlopen(request, timeout=REQUEST_TIMEOUT) as response:
        content = response.read(1024 * 1024 + 1)
    if len(content) > 1024 * 1024:
        raise UpdateError("The release checksum file is unexpectedly large.")
    try:
        lines = content.decode("utf-8").splitlines()
    except UnicodeDecodeError as error:
        raise UpdateError("The release checksum file is not valid UTF-8.") from error

    for line in lines:
        parts = line.split()
        if len(parts) == 2 and parts[1].lstrip("*") == EXECUTABLE_NAME:
            checksum = parts[0].lower()
            if re.fullmatch(r"[0-9a-f]{64}", checksum):
                return checksum
    raise UpdateError(f"No valid checksum was found for {EXECUTABLE_NAME}.")


def download_release(release):
    downloaded_path = None

    def remove_partial_download():
        if downloaded_path is None:
            return
        try:
            downloaded_path.unlink(missing_ok=True)
        except OSError as error:
            raise UpdateError(f"Could not remove the incomplete update: {error}") from error

    try:
        expected_hash = _read_checksums(release.checksum_url)
        request = Request(
            release.executable_url,
            headers={"User-Agent": "AlumetalManager-Updater"},
        )
        with urlopen(request, timeout=REQUEST_TIMEOUT) as response:
            with tempfile.NamedTemporaryFile(
                prefix="AlumetalManager-update-",
                suffix=".exe",
                delete=False,
            ) as output:
                downloaded_path = Path(output.name)
                digest = hashlib.sha256()
                size = 0
                while True:
                    chunk = response.read(1024 * 1024)
                    if not chunk:
                        break
                    size += len(chunk)
                    if size > release.executable_size:
                        raise UpdateError("The downloaded update is larger than expected.")
                    digest.update(chunk)
                    output.write(chunk)

        if size != release.executable_size:
            raise UpdateError("The downloaded update size does not match the release.")
        if digest.hexdigest() != expected_hash:
            raise UpdateError("The downloaded update failed its SHA-256 checksum.")
        return downloaded_path
    except UpdateError:
        remove_partial_download()
        raise
    except (HTTPError, URLError, TimeoutError, http.client.HTTPException, OSError, ValueError) as error:
        remove_partial_download()
        raise UpdateError(f"Could not download the update: {error}") from error


def _update_directory():
    root = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    directory = Path(root) / "AlumetalManager" / "updates"
    try:
        directory.mkdir(parents=True, exist_ok=True)
    except OSError as error:
        raise UpdateError(f"Could not create the update folder: {error}") from error
    return directory


def start_windows_update(downloaded_executable, current_executable):
    if os.name != "nt":
        raise UpdateError("Automatic installation is supported only on Windows.")
    powershell = shutil.which("powershell.exe")
    if not powershell:
        raise UpdateError("Windows PowerShell could not be found.")

    target_path = Path(current_executable).resolve()
    downloaded_path = Path(downloaded_executable).resolve()
    if not target_path.is_file() or not downloaded_path.is_file():
        raise UpdateError("The current program or downloaded update is missing.")

    directory = _update_directory()
    script_path = directory / f"install-{uuid.uuid4().hex}.ps1"
    log_path = directory / "update-error.log"
    script = r"""
param(
    [int]$ProcessId,
    [string]$TargetPath,
    [string]$DownloadedPath,
    [string]$LogPath
)
$ErrorActionPreference = 'Stop'
$replacementPath = "$TargetPath.update"
$backupPath = "$TargetPath.previous"
try {
    Wait-Process -Id $ProcessId -ErrorAction SilentlyContinue
    Copy-Item -LiteralPath $DownloadedPath -Destination $replacementPath -Force
    if (Test-Path -LiteralPath $TargetPath) {
        Move-Item -LiteralPath $TargetPath -Destination $backupPath -Force
    }
    try {
        Move-Item -LiteralPath $replacementPath -Destination $TargetPath -Force
    } catch {
        if (Test-Path -LiteralPath $backupPath) {
            Move-Item -LiteralPath $backupPath -Destination $TargetPath -Force
        }
        throw
    }
    if (Test-Path -LiteralPath $backupPath) {
        Remove-Item -LiteralPath $backupPath -Force
    }
    Start-Process -FilePath $TargetPath
} catch {
    Set-Content -LiteralPath $LogPath -Value $_.Exception.Message -Encoding UTF8
    try {
        Add-Type -AssemblyName System.Windows.Forms
        [System.Windows.Forms.MessageBox]::Show(
            "The update could not be installed. Details: $LogPath",
            "Alumetal Manager",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    } catch {}
} finally {
    Remove-Item -LiteralPath $DownloadedPath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $replacementPath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue
}
"""
    try:
        script_path.write_text(script, encoding="utf-8")
        subprocess.Popen(
            [
                powershell,
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-WindowStyle",
                "Hidden",
                "-File",
                str(script_path),
                "-ProcessId",
                str(os.getpid()),
                "-TargetPath",
                str(target_path),
                "-DownloadedPath",
                str(downloaded_path),
                "-LogPath",
                str(log_path),
            ],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            close_fds=True,
            creationflags=(
                subprocess.DETACHED_PROCESS
                | subprocess.CREATE_NEW_PROCESS_GROUP
                | subprocess.CREATE_NO_WINDOW
            ),
        )
    except (OSError, ValueError) as error:
        script_path.unlink(missing_ok=True)
        raise UpdateError(f"Could not start the update installer: {error}") from error
    return log_path
