"""OfficeCLI 固定运行时解析与缓存校验。

该模块实现 v5 规范中的 V5-001、V5-002 地基能力：
- 读取并校验 `officecli.lock.json`
- 按系统、架构和 libc 解析 runtime_id
- 按锁文件下载、缓存、哈希和版本校验 OfficeCLI
- stdout 仅输出 JSON，诊断信息由调用方记录
"""

from __future__ import annotations

import argparse
from contextlib import contextmanager
import hashlib
import json
import os
import platform
import re
import socket
import stat
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
import uuid
from pathlib import Path
from typing import Any, Dict, Iterable, Optional


FH_OFFICECLI_LOCK_INVALID = "FH-OFFICECLI-LOCK-INVALID"
FH_OFFICECLI_PLATFORM_UNSUPPORTED = "FH-OFFICECLI-PLATFORM-UNSUPPORTED"
FH_OFFICECLI_DOWNLOAD_FAILED = "FH-OFFICECLI-DOWNLOAD-FAILED"
FH_OFFICECLI_HASH_MISMATCH = "FH-OFFICECLI-HASH-MISMATCH"
FH_OFFICECLI_VERSION_MISMATCH = "FH-OFFICECLI-VERSION-MISMATCH"
FH_OFFICECLI_OFFLINE_CACHE_MISS = "FH-OFFICECLI-OFFLINE-CACHE-MISS"
FH_OFFICECLI_CACHE_INVALID = "FH-OFFICECLI-CACHE-INVALID"


EXPECTED_RUNTIME_IDS = {
    "win-x64",
    "win-arm64",
    "linux-x64-gnu",
    "linux-arm64-gnu",
    "linux-x64-musl",
    "linux-arm64-musl",
    "osx-x64",
    "osx-arm64",
}


EXPECTED_ASSETS = {
    "win-x64": {
        "os": "windows",
        "arch": "x64",
        "libc": None,
        "asset_name": "officecli-win-x64.exe",
        "sha256": "15d29f3a04e6ad00503de178f98dae872b47ef71f09fac3c614212b209c4d229",
        "size_bytes": 31997816,
        "primary_url": "https://github.com/iOfficeAI/OfficeCLI/releases/download/v1.0.113/officecli-win-x64.exe",
        "mirror_url": None,
        "executable_name": "officecli.exe",
    },
    "win-arm64": {
        "os": "windows",
        "arch": "arm64",
        "libc": None,
        "asset_name": "officecli-win-arm64.exe",
        "sha256": "94fa5101b94f2fe59c1458688bbc3ddcde4f244afe204143b7eac9bb5089f784",
        "size_bytes": 32448388,
        "primary_url": "https://github.com/iOfficeAI/OfficeCLI/releases/download/v1.0.113/officecli-win-arm64.exe",
        "mirror_url": None,
        "executable_name": "officecli.exe",
    },
    "linux-x64-gnu": {
        "os": "linux",
        "arch": "x64",
        "libc": "glibc",
        "asset_name": "officecli-linux-x64",
        "sha256": "ffe09f5f8ec76240e44ff431b802b8a4466775afda328f1f7b606e3a79807311",
        "size_bytes": 33950776,
        "primary_url": "https://github.com/iOfficeAI/OfficeCLI/releases/download/v1.0.113/officecli-linux-x64",
        "mirror_url": None,
        "executable_name": "officecli",
    },
    "linux-arm64-gnu": {
        "os": "linux",
        "arch": "arm64",
        "libc": "glibc",
        "asset_name": "officecli-linux-arm64",
        "sha256": "893874471e6830d29580ba9cab0a5834eab80278092f77edb31292bffff1f9fd",
        "size_bytes": 33369562,
        "primary_url": "https://github.com/iOfficeAI/OfficeCLI/releases/download/v1.0.113/officecli-linux-arm64",
        "mirror_url": None,
        "executable_name": "officecli",
    },
    "linux-x64-musl": {
        "os": "linux",
        "arch": "x64",
        "libc": "musl",
        "asset_name": "officecli-linux-alpine-x64",
        "sha256": "5579d760de781781c7a05e32774bea0bdd091ad3ba3d013129a35e2c837a09be",
        "size_bytes": 33968418,
        "primary_url": "https://github.com/iOfficeAI/OfficeCLI/releases/download/v1.0.113/officecli-linux-alpine-x64",
        "mirror_url": None,
        "executable_name": "officecli",
    },
    "linux-arm64-musl": {
        "os": "linux",
        "arch": "arm64",
        "libc": "musl",
        "asset_name": "officecli-linux-alpine-arm64",
        "sha256": "a18f81e2a4f9cbc8bbec80fc305b20aec1352327094bdff1b48fdc13da3dddba",
        "size_bytes": 33410098,
        "primary_url": "https://github.com/iOfficeAI/OfficeCLI/releases/download/v1.0.113/officecli-linux-alpine-arm64",
        "mirror_url": None,
        "executable_name": "officecli",
    },
    "osx-x64": {
        "os": "macos",
        "arch": "x64",
        "libc": None,
        "asset_name": "officecli-mac-x64",
        "sha256": "62ad1b63ec1b833efe01a51d3564238ce274b51a785b1a2fc91880c66381b0d2",
        "size_bytes": 33330224,
        "primary_url": "https://github.com/iOfficeAI/OfficeCLI/releases/download/v1.0.113/officecli-mac-x64",
        "mirror_url": None,
        "executable_name": "officecli",
    },
    "osx-arm64": {
        "os": "macos",
        "arch": "arm64",
        "libc": None,
        "asset_name": "officecli-mac-arm64",
        "sha256": "35a733b598cb32a57d4edc1217a5edfcf63aa9c141916b0b4ef54aa37e4c30ba",
        "size_bytes": 32587584,
        "primary_url": "https://github.com/iOfficeAI/OfficeCLI/releases/download/v1.0.113/officecli-mac-arm64",
        "mirror_url": None,
        "executable_name": "officecli",
    },
}


class OfficeCliRuntimeError(Exception):
    """OfficeCLI 运行时解析错误。"""

    def __init__(self, code: str, message: str, detail: Optional[dict] = None):
        self.code = code
        self.message = message
        self.detail = detail or {}
        super().__init__(f"[{code}] {message}")

    def to_json(self) -> dict:
        """转换为稳定 JSON 错误对象。"""
        return {
            "ok": False,
            "error": {
                "code": self.code,
                "message": self.message,
                "detail": self.detail,
            },
        }


def read_json(path: Path) -> dict:
    """读取 UTF-8 JSON 文件。"""
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_LOCK_INVALID,
            f"读取锁文件失败：{path}",
            {"reason": str(exc)},
        )


def load_lock(lock_path: Path) -> dict:
    """读取并校验 OfficeCLI 锁文件。"""
    lock = read_json(lock_path)
    validate_lock(lock)
    return lock


def _require_keys(obj: dict, keys: Iterable[str], path: str) -> None:
    missing = [key for key in keys if key not in obj]
    if missing:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_LOCK_INVALID,
            f"{path} 缺少必填字段",
            {"missing": missing},
        )


def validate_lock(lock: dict) -> None:
    """执行锁文件结构和固定值校验。"""
    _require_keys(
        lock,
        [
            "schema_id",
            "schema_version",
            "officecli_version",
            "release_tag",
            "source_commit",
            "released_at",
            "license",
            "primary_base_url",
            "mirror_base_url",
            "auto_update_disabled",
            "assets",
        ],
        "$",
    )
    expected = {
        "schema_id": "officecli-lock",
        "schema_version": "1.0.0",
        "officecli_version": "1.0.113",
        "release_tag": "v1.0.113",
        "source_commit": "8e5b17977493de1b46536561a50971799c4fc665",
        "license": "Apache-2.0",
    }
    drift = {
        key: {"expected": value, "actual": lock.get(key)}
        for key, value in expected.items()
        if lock.get(key) != value
    }
    if drift:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_LOCK_INVALID,
            "锁文件固定字段漂移",
            {"drift": drift},
        )
    if lock["mirror_base_url"] is not None or lock["auto_update_disabled"] is not True:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_LOCK_INVALID,
            "锁文件镜像或自动更新策略不符合 v5 固定契约",
        )
    assets = lock["assets"]
    if not isinstance(assets, list) or len(assets) != 8:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_LOCK_INVALID,
            "锁文件必须精确包含 8 个平台资产",
        )
    seen = set()
    for index, asset in enumerate(assets):
        _require_keys(
            asset,
            [
                "runtime_id",
                "os",
                "arch",
                "libc",
                "asset_name",
                "sha256",
                "size_bytes",
                "primary_url",
                "mirror_url",
                "executable_name",
            ],
            f"$.assets[{index}]",
        )
        runtime_id = asset["runtime_id"]
        if runtime_id not in EXPECTED_RUNTIME_IDS:
            raise OfficeCliRuntimeError(
                FH_OFFICECLI_LOCK_INVALID,
                f"未知 runtime_id：{runtime_id}",
            )
        if runtime_id in seen:
            raise OfficeCliRuntimeError(
                FH_OFFICECLI_LOCK_INVALID,
                f"重复 runtime_id：{runtime_id}",
            )
        seen.add(runtime_id)
        expected_asset = EXPECTED_ASSETS[runtime_id]
        asset_drift = {
            key: {"expected": value, "actual": asset.get(key)}
            for key, value in expected_asset.items()
            if asset.get(key) != value
        }
        if asset_drift:
            raise OfficeCliRuntimeError(
                FH_OFFICECLI_LOCK_INVALID,
                f"{runtime_id} 的官方资产元组漂移",
                {"runtime_id": runtime_id, "drift": asset_drift},
            )
        if len(asset["sha256"]) != 64 or asset["sha256"].lower() != asset["sha256"]:
            raise OfficeCliRuntimeError(
                FH_OFFICECLI_LOCK_INVALID,
                f"{runtime_id} 的 sha256 格式不合法",
            )
        if asset["mirror_url"] is not None:
            raise OfficeCliRuntimeError(
                FH_OFFICECLI_LOCK_INVALID,
                f"{runtime_id} 首版不允许配置 mirror_url",
            )
        if not asset["primary_url"].endswith("/" + asset["asset_name"]):
            raise OfficeCliRuntimeError(
                FH_OFFICECLI_LOCK_INVALID,
                f"{runtime_id} 的 primary_url 与 asset_name 不一致",
            )
    if seen != EXPECTED_RUNTIME_IDS:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_LOCK_INVALID,
            "锁文件 runtime_id 集合不完整",
            {"missing": sorted(EXPECTED_RUNTIME_IDS - seen), "extra": sorted(seen - EXPECTED_RUNTIME_IDS)},
        )


def normalize_arch(machine: str) -> Optional[str]:
    """归一化 CPU 架构。"""
    value = machine.lower()
    if value in {"amd64", "x86_64", "x64"}:
        return "x64"
    if value in {"arm64", "aarch64"}:
        return "arm64"
    return None


def detect_linux_libc(
    alpine_release_exists: Optional[bool] = None,
    ldd_output: Optional[str] = None,
) -> str:
    """识别 Linux libc 类型。"""
    if alpine_release_exists is None:
        alpine_release_exists = Path("/etc/alpine-release").exists()
    if alpine_release_exists:
        return "musl"
    if ldd_output is None:
        try:
            proc = subprocess.run(
                ["ldd", "--version"],
                text=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                timeout=5,
                check=False,
            )
            ldd_output = (proc.stdout or "") + "\n" + (proc.stderr or "")
        except (OSError, subprocess.TimeoutExpired):
            ldd_output = ""
    if "musl" in ldd_output.lower():
        return "musl"
    return "glibc"


def detect_runtime_id(
    system_name: Optional[str] = None,
    machine: Optional[str] = None,
    alpine_release_exists: Optional[bool] = None,
    ldd_output: Optional[str] = None,
) -> str:
    """按 v5 固定映射解析 runtime_id。"""
    system = (system_name or platform.system()).lower()
    arch = normalize_arch(machine or platform.machine())
    if arch is None:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_PLATFORM_UNSUPPORTED,
            "不支持的 CPU 架构",
            {"machine": machine or platform.machine()},
        )
    if system in {"windows", "win32"}:
        return f"win-{arch}"
    if system == "darwin":
        return f"osx-{arch}"
    if system == "linux":
        libc = detect_linux_libc(alpine_release_exists, ldd_output)
        return f"linux-{arch}-{'musl' if libc == 'musl' else 'gnu'}"
    raise OfficeCliRuntimeError(
        FH_OFFICECLI_PLATFORM_UNSUPPORTED,
        "不支持的操作系统",
        {"system": system_name or platform.system()},
    )


def select_asset(lock: dict, runtime_id: str) -> dict:
    """从锁文件中选择 runtime_id 对应资产。"""
    for asset in lock["assets"]:
        if asset["runtime_id"] == runtime_id:
            return asset
    raise OfficeCliRuntimeError(
        FH_OFFICECLI_PLATFORM_UNSUPPORTED,
        f"锁文件缺少 runtime_id：{runtime_id}",
    )


def sha256_file(path: Path) -> str:
    """计算文件 SHA-256。"""
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def cache_executable_path(workspace_root: Path, lock: dict, asset: dict) -> Path:
    """返回固定缓存可执行文件路径。"""
    return (
        workspace_root
        / ".cache"
        / "officecli"
        / f"v{lock['officecli_version']}"
        / asset["runtime_id"]
        / asset["executable_name"]
    )


def verify_file_hash_and_size(path: Path, asset: dict) -> dict:
    """校验缓存文件大小和 SHA-256。"""
    if not path.exists():
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_OFFLINE_CACHE_MISS,
            "OfficeCLI 缓存不存在",
            {"path": str(path)},
        )
    size = path.stat().st_size
    if size != asset["size_bytes"]:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_CACHE_INVALID,
            "OfficeCLI 缓存大小与锁文件不一致",
            {"path": str(path), "expected": asset["size_bytes"], "actual": size},
        )
    actual_sha256 = sha256_file(path)
    if actual_sha256 != asset["sha256"]:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_HASH_MISMATCH,
            "OfficeCLI 缓存 SHA-256 与锁文件不一致",
            {"path": str(path), "expected": asset["sha256"], "actual": actual_sha256},
        )
    return {"size_bytes": size, "sha256": actual_sha256}


def ensure_executable_permission(path: Path) -> None:
    """在 POSIX 平台确保所有者执行位存在。"""
    if os.name == "nt":
        return
    current = path.stat().st_mode
    if not current & stat.S_IXUSR:
        path.chmod(current | stat.S_IXUSR)


def run_version_check(path: Path, expected_version: str, timeout_seconds: int = 20) -> str:
    """执行 `officecli --version` 并精确校验版本。"""
    env = os.environ.copy()
    env["OFFICECLI_SKIP_UPDATE"] = "1"
    env["OFFICECLI_NO_AUTO_RESIDENT"] = "1"
    try:
        proc = subprocess.run(
            [str(path), "--version"],
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout_seconds,
            check=False,
            env=env,
        )
    except (OSError, subprocess.TimeoutExpired) as exc:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_VERSION_MISMATCH,
            "OfficeCLI 版本检查无法执行",
            {"path": str(path), "reason": str(exc)},
        )
    combined = ((proc.stdout or "") + "\n" + (proc.stderr or "")).strip()
    if proc.returncode != 0 or not is_exact_version_output(combined, expected_version):
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_VERSION_MISMATCH,
            "OfficeCLI 运行版本与锁文件不一致",
            {
                "path": str(path),
                "expected": expected_version,
                "exit_code": proc.returncode,
                "output": combined,
            },
        )
    return expected_version


def is_exact_version_output(output: str, expected_version: str) -> bool:
    """严格校验 `officecli --version` 输出版本。"""
    lines = [line.strip() for line in output.splitlines() if line.strip()]
    if len(lines) != 1:
        return False
    patterns = [
        rf"^{re.escape(expected_version)}$",
        rf"^v{re.escape(expected_version)}$",
        rf"^OfficeCLI\s+{re.escape(expected_version)}$",
        rf"^OfficeCLI\s+v{re.escape(expected_version)}$",
    ]
    return any(re.fullmatch(pattern, lines[0], flags=re.IGNORECASE) for pattern in patterns)


def download_asset(asset: dict, destination: Path, timeout_seconds: int = 120) -> None:
    """从固定 GitHub Release URL 下载资产到临时路径。"""
    try:
        with urllib.request.urlopen(asset["primary_url"], timeout=timeout_seconds) as response:
            with destination.open("wb") as handle:
                while True:
                    chunk = response.read(1024 * 1024)
                    if not chunk:
                        break
                    handle.write(chunk)
    except (OSError, urllib.error.URLError) as exc:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_DOWNLOAD_FAILED,
            "OfficeCLI 资产下载失败",
            {"url": asset["primary_url"], "reason": str(exc)},
        )


def is_windows_file_in_use_error(exc: OSError) -> bool:
    """判断是否为 Windows 目标文件被占用错误。"""
    return isinstance(exc, PermissionError) and getattr(exc, "winerror", None) == 32


def install_downloaded_asset(temp_path: Path, target: Path, asset: dict, file_info: dict) -> dict:
    """安装已校验下载文件；若 Windows 目标被占用但合法，则复用目标缓存。"""
    try:
        temp_path.replace(target)
        return {"status": "downloaded", **file_info}
    except OSError as exc:
        if not is_windows_file_in_use_error(exc):
            raise OfficeCliRuntimeError(
                FH_OFFICECLI_DOWNLOAD_FAILED,
                "OfficeCLI 缓存替换失败",
                {"target": str(target), "temp_path": str(temp_path), "reason": str(exc)},
            ) from exc
        try:
            existing_info = verify_file_hash_and_size(target, asset)
        except OfficeCliRuntimeError:
            raise OfficeCliRuntimeError(
                FH_OFFICECLI_DOWNLOAD_FAILED,
                "OfficeCLI 缓存替换失败",
                {"target": str(target), "temp_path": str(temp_path), "reason": str(exc)},
            ) from exc
        ensure_executable_permission(target)
        return {"status": "cached", **existing_info}


def process_exists(pid: int) -> bool:
    """判断进程是否仍存在。"""
    if pid <= 0:
        return False
    try:
        os.kill(pid, 0)
        return True
    except OSError:
        return False


def append_lock_audit(lock_dir: Path, event: dict) -> None:
    """记录运行时锁审计事件。"""
    audit_path = lock_dir / "runtime-lock.audit.jsonl"
    event_with_time = {
        "recorded_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        **event,
    }
    with audit_path.open("a", encoding="utf-8", newline="\n") as handle:
        handle.write(json.dumps(event_with_time, ensure_ascii=False, sort_keys=True) + "\n")


@contextmanager
def runtime_file_lock(lock_dir: Path, runtime_id: str, wait_seconds: int = 60):
    """获取 `{runtime_id}.lock` 排他文件锁。"""
    lock_dir.mkdir(parents=True, exist_ok=True)
    lock_path = lock_dir / f"{runtime_id}.lock"
    nonce = uuid.uuid4().hex
    started = time.time()
    payload = {
        "pid": os.getpid(),
        "host": socket.gethostname(),
        "started_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(started)),
        "nonce": nonce,
    }
    while True:
        try:
            fd = os.open(str(lock_path), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            with os.fdopen(fd, "w", encoding="utf-8", newline="\n") as handle:
                json.dump(payload, handle, ensure_ascii=False, sort_keys=True)
                handle.write("\n")
            break
        except FileExistsError:
            now = time.time()
            stale = False
            stale_reason = None
            try:
                current = json.loads(lock_path.read_text(encoding="utf-8"))
                pid = int(current.get("pid", -1))
                lock_age = now - lock_path.stat().st_mtime
                stale = lock_age > 600 and not process_exists(pid)
                stale_reason = {"pid": pid, "lock_age_seconds": round(lock_age, 3)}
            except (OSError, ValueError, json.JSONDecodeError) as exc:
                lock_age = now - lock_path.stat().st_mtime if lock_path.exists() else 0
                stale = lock_age > 600
                stale_reason = {"reason": str(exc), "lock_age_seconds": round(lock_age, 3)}
            if stale:
                append_lock_audit(
                    lock_dir,
                    {
                        "event": "remove_stale_lock",
                        "runtime_id": runtime_id,
                        "lock_path": str(lock_path),
                        "detail": stale_reason,
                    },
                )
                try:
                    lock_path.unlink()
                except FileNotFoundError:
                    pass
                continue
            if now - started >= wait_seconds:
                raise OfficeCliRuntimeError(
                    FH_OFFICECLI_CACHE_INVALID,
                    "等待 OfficeCLI runtime 文件锁超时",
                    {"runtime_id": runtime_id, "lock_path": str(lock_path), "wait_seconds": wait_seconds},
                )
            time.sleep(0.2)
    try:
        yield lock_path
    finally:
        try:
            current = json.loads(lock_path.read_text(encoding="utf-8"))
            if current.get("nonce") == nonce:
                lock_path.unlink()
        except FileNotFoundError:
            pass


def materialize_asset(
    lock: dict,
    asset: dict,
    target: Path,
    offline: bool,
    skip_version_check: bool,
) -> dict:
    """确保目标缓存存在且通过校验。"""
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        file_info = verify_file_hash_and_size(target, asset)
        ensure_executable_permission(target)
        version = None if skip_version_check else run_version_check(target, lock["officecli_version"])
        return {"status": "cached", "version": version or lock["officecli_version"], **file_info}
    if offline:
        raise OfficeCliRuntimeError(
            FH_OFFICECLI_OFFLINE_CACHE_MISS,
            "无网络模式下 OfficeCLI 缓存不存在",
            {"path": str(target), "runtime_id": asset["runtime_id"]},
        )
    with runtime_file_lock(target.parent, asset["runtime_id"]):
        if target.exists():
            file_info = verify_file_hash_and_size(target, asset)
            ensure_executable_permission(target)
            version = None if skip_version_check else run_version_check(target, lock["officecli_version"])
            return {"status": "cached", "version": version or lock["officecli_version"], **file_info}
        temp_dir = target.parent
        temp_dir.mkdir(parents=True, exist_ok=True)
        handle, temp_name = tempfile.mkstemp(
            prefix=f".download.{os.getpid()}.{uuid.uuid4().hex}.",
            suffix=".tmp",
            dir=str(temp_dir),
        )
        os.close(handle)
        temp_path = Path(temp_name)
        try:
            download_asset(asset, temp_path)
            file_info = verify_file_hash_and_size(temp_path, asset)
            ensure_executable_permission(temp_path)
            version = None if skip_version_check else run_version_check(temp_path, lock["officecli_version"])
            installed = install_downloaded_asset(temp_path, target, asset, file_info)
            return {"version": version or lock["officecli_version"], **installed}
        finally:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except OSError:
                    pass


def ensure_officecli(
    lock_path: Path,
    workspace_root: Optional[Path] = None,
    runtime_id: Optional[str] = None,
    offline: bool = False,
    skip_version_check: bool = False,
) -> dict:
    """解析并确保固定 OfficeCLI 可执行文件可用。"""
    lock_path = Path(lock_path).resolve()
    lock = load_lock(lock_path)
    resolved_workspace = Path(workspace_root).resolve() if workspace_root else lock_path.parent.parent.parent.resolve()
    resolved_runtime_id = runtime_id or detect_runtime_id()
    asset = select_asset(lock, resolved_runtime_id)
    executable = cache_executable_path(resolved_workspace, lock, asset)
    materialized = materialize_asset(lock, asset, executable, offline, skip_version_check)
    return {
        "ok": True,
        "schema_id": "officecli-runtime-resolution",
        "schema_version": "1.0.0",
        "officecli_version": lock["officecli_version"],
        "runtime_id": resolved_runtime_id,
        "asset_name": asset["asset_name"],
        "executable_path": str(executable),
        "cache_status": materialized["status"],
        "sha256": materialized["sha256"],
        "size_bytes": materialized["size_bytes"],
        "version": materialized["version"],
    }


def parse_args(argv: Optional[list] = None) -> argparse.Namespace:
    """解析命令行参数。"""
    parser = argparse.ArgumentParser(description="OfficeCLI 固定运行时解析器")
    sub = parser.add_subparsers(dest="command", required=True)
    ensure = sub.add_parser("ensure", help="获取并校验固定 OfficeCLI")
    ensure.add_argument("--lock", required=True, type=Path)
    ensure.add_argument("--workspace-root", type=Path)
    ensure.add_argument("--runtime-id")
    ensure.add_argument("--offline", action="store_true")
    return parser.parse_args(argv)


def main(argv: Optional[list] = None) -> int:
    """CLI 入口。"""
    args = parse_args(argv)
    try:
        if args.command == "ensure":
            result = ensure_officecli(
                lock_path=args.lock,
                workspace_root=args.workspace_root,
                runtime_id=args.runtime_id,
                offline=args.offline,
                skip_version_check=False,
            )
        else:
            raise OfficeCliRuntimeError(FH_OFFICECLI_LOCK_INVALID, f"未知命令：{args.command}")
        sys.stdout.write(json.dumps(result, ensure_ascii=False, sort_keys=True) + "\n")
        return 0
    except OfficeCliRuntimeError as exc:
        sys.stdout.write(json.dumps(exc.to_json(), ensure_ascii=False, sort_keys=True) + "\n")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
