"""OfficeCLI 运行时解析器测试。

覆盖 OFFICECLI-T001 到 OFFICECLI-T014 中无需真实下载即可验证的供应链和平台规则。
"""

from __future__ import annotations

import hashlib
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from scripts.officecli.runtime_resolver import (  # noqa: E402
    EXPECTED_RUNTIME_IDS,
    FH_OFFICECLI_CACHE_INVALID,
    FH_OFFICECLI_DOWNLOAD_FAILED,
    FH_OFFICECLI_OFFLINE_CACHE_MISS,
    FH_OFFICECLI_PLATFORM_UNSUPPORTED,
    OfficeCliRuntimeError,
    cache_executable_path,
    detect_runtime_id,
    is_exact_version_output,
    install_downloaded_asset,
    load_lock,
    materialize_asset,
    parse_args,
    runtime_file_lock,
    select_asset,
    validate_lock,
    verify_file_hash_and_size,
)


LOCK_PATH = Path(__file__).parent.parent.parent / "tools" / "officecli" / "officecli.lock.json"


class FakeWindowsFileInUseError(PermissionError):
    """测试用 WinError 32 异常，避免依赖不同平台的 OSError 构造语义。"""

    winerror = 32


class TestOfficeCliLock(unittest.TestCase):
    """测试 OfficeCLI 锁文件。"""

    def test_lock_contains_exact_eight_assets(self):
        """OFFICECLI-T001：锁文件必须精确包含八个平台资产。"""
        lock = load_lock(LOCK_PATH)
        runtime_ids = {asset["runtime_id"] for asset in lock["assets"]}

        self.assertEqual(lock["officecli_version"], "1.0.113")
        self.assertEqual(lock["release_tag"], "v1.0.113")
        self.assertEqual(runtime_ids, EXPECTED_RUNTIME_IDS)
        self.assertEqual(len(lock["assets"]), 8)
        self.assertTrue(lock["auto_update_disabled"])

    def test_lock_drift_is_rejected(self):
        """OFFICECLI-T006：锁文件版本漂移必须阻塞。"""
        lock = load_lock(LOCK_PATH)
        lock["officecli_version"] = "1.0.114"

        with self.assertRaises(OfficeCliRuntimeError) as ctx:
            validate_lock(lock)

        self.assertEqual(ctx.exception.code, "FH-OFFICECLI-LOCK-INVALID")

    def test_lock_asset_tuple_drift_is_rejected(self):
        """OFFICECLI-T006：任一平台官方资产元组漂移必须阻塞。"""
        lock = load_lock(LOCK_PATH)
        lock["assets"][0]["sha256"] = "0" * 64

        with self.assertRaises(OfficeCliRuntimeError) as ctx:
            validate_lock(lock)

        self.assertEqual(ctx.exception.code, "FH-OFFICECLI-LOCK-INVALID")
        self.assertIn("runtime_id", ctx.exception.detail)

    def test_select_asset(self):
        """正向：可按 runtime_id 选择唯一资产。"""
        lock = load_lock(LOCK_PATH)
        asset = select_asset(lock, "linux-x64-musl")

        self.assertEqual(asset["asset_name"], "officecli-linux-alpine-x64")
        self.assertEqual(asset["libc"], "musl")


class TestRuntimeDetection(unittest.TestCase):
    """测试平台映射。"""

    def test_detect_windows_x64(self):
        """OFFICECLI-T007：Windows x64 映射到 win-x64。"""
        self.assertEqual(detect_runtime_id("Windows", "AMD64"), "win-x64")

    def test_detect_macos_arm64(self):
        """OFFICECLI-T008：macOS arm64 映射到 osx-arm64。"""
        self.assertEqual(detect_runtime_id("Darwin", "arm64"), "osx-arm64")

    def test_detect_linux_glibc(self):
        """OFFICECLI-T009：Linux glibc 映射到 gnu 资产。"""
        result = detect_runtime_id(
            "Linux",
            "x86_64",
            alpine_release_exists=False,
            ldd_output="ldd (GNU libc) 2.39",
        )

        self.assertEqual(result, "linux-x64-gnu")

    def test_detect_linux_musl(self):
        """OFFICECLI-T010：Linux musl/Alpine 映射到 alpine 资产。"""
        result = detect_runtime_id(
            "Linux",
            "aarch64",
            alpine_release_exists=True,
            ldd_output="",
        )

        self.assertEqual(result, "linux-arm64-musl")

    def test_unsupported_arch_is_rejected(self):
        """OFFICECLI-T011：未知架构不得猜测。"""
        with self.assertRaises(OfficeCliRuntimeError) as ctx:
            detect_runtime_id("Linux", "riscv64")

        self.assertEqual(ctx.exception.code, FH_OFFICECLI_PLATFORM_UNSUPPORTED)


class TestCacheValidation(unittest.TestCase):
    """测试缓存路径和文件校验。"""

    def test_cache_path_uses_locked_version_and_runtime(self):
        """OFFICECLI-T002：缓存路径必须使用固定版本和 runtime_id。"""
        lock = load_lock(LOCK_PATH)
        asset = select_asset(lock, "win-x64")
        with tempfile.TemporaryDirectory() as tmp:
            path = cache_executable_path(Path(tmp), lock, asset)

        self.assertTrue(str(path).replace("\\", "/").endswith(".cache/officecli/v1.0.113/win-x64/officecli.exe"))

    def test_missing_offline_cache_is_rejected(self):
        """OFFICECLI-T013：无网且缓存不存在必须阻塞。"""
        lock = load_lock(LOCK_PATH)
        asset = select_asset(lock, "win-x64")
        with tempfile.TemporaryDirectory() as tmp:
            missing = Path(tmp) / "officecli.exe"
            with self.assertRaises(OfficeCliRuntimeError) as ctx:
                verify_file_hash_and_size(missing, asset)

        self.assertEqual(ctx.exception.code, FH_OFFICECLI_OFFLINE_CACHE_MISS)

    def test_cache_size_mismatch_is_rejected(self):
        """OFFICECLI-T014：缓存损坏必须阻塞。"""
        fake_asset = {
            "size_bytes": 10,
            "sha256": hashlib.sha256(b"abc").hexdigest(),
        }
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "officecli"
            path.write_bytes(b"abc")
            with self.assertRaises(OfficeCliRuntimeError) as ctx:
                verify_file_hash_and_size(path, fake_asset)

        self.assertEqual(ctx.exception.code, FH_OFFICECLI_CACHE_INVALID)

    def test_cache_hash_match(self):
        """正向：大小和 SHA-256 匹配时返回文件证据。"""
        content = b"abc"
        fake_asset = {
            "size_bytes": len(content),
            "sha256": hashlib.sha256(content).hexdigest(),
        }
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "officecli"
            path.write_bytes(content)
            result = verify_file_hash_and_size(path, fake_asset)

        self.assertEqual(result["size_bytes"], len(content))
        self.assertEqual(result["sha256"], fake_asset["sha256"])

    def test_install_downloaded_asset_reuses_valid_locked_target(self):
        """Windows 目标文件被占用但缓存合法时必须复用，避免误报下载失败。"""
        content = b"officecli"
        fake_asset = {
            "runtime_id": "win-x64",
            "size_bytes": len(content),
            "sha256": hashlib.sha256(content).hexdigest(),
        }
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "officecli.exe"
            temp_path = Path(tmp) / ".download.tmp"
            target.write_bytes(content)
            temp_path.write_bytes(content)
            file_info = verify_file_hash_and_size(temp_path, fake_asset)
            replace_error = FakeWindowsFileInUseError(13, "文件正由另一进程使用", str(target))

            with mock.patch(
                "scripts.officecli.runtime_resolver.Path.replace",
                side_effect=replace_error,
            ):
                result = install_downloaded_asset(temp_path, target, fake_asset, file_info)

            self.assertEqual(result["status"], "cached")
            self.assertEqual(result["size_bytes"], len(content))
            self.assertEqual(result["sha256"], fake_asset["sha256"])

    def test_install_downloaded_asset_blocks_unexpected_replace_error(self):
        """非 WinError 32 的替换失败不得被合法旧缓存掩盖。"""
        content = b"officecli"
        fake_asset = {
            "runtime_id": "win-x64",
            "size_bytes": len(content),
            "sha256": hashlib.sha256(content).hexdigest(),
        }
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "officecli.exe"
            temp_path = Path(tmp) / ".download.tmp"
            target.write_bytes(content)
            temp_path.write_bytes(content)
            file_info = verify_file_hash_and_size(temp_path, fake_asset)

            with mock.patch(
                "scripts.officecli.runtime_resolver.Path.replace",
                side_effect=OSError(13, "权限被拒绝", str(target), 5),
            ):
                with self.assertRaises(OfficeCliRuntimeError) as ctx:
                    install_downloaded_asset(temp_path, target, fake_asset, file_info)

            self.assertEqual(ctx.exception.code, FH_OFFICECLI_DOWNLOAD_FAILED)

    def test_materialize_asset_cleans_temp_when_reusing_locked_target(self):
        """下载调用链复用被占用目标缓存后仍必须清理当前临时文件。"""
        content = b"officecli"
        fake_asset = {
            "runtime_id": "win-x64",
            "size_bytes": len(content),
            "sha256": hashlib.sha256(content).hexdigest(),
        }
        fake_lock = {"officecli_version": "1.0.113"}
        with tempfile.TemporaryDirectory() as tmp:
            target = Path(tmp) / "officecli.exe"

            def fake_download(_asset, destination):
                destination.write_bytes(content)
                target.write_bytes(content)

            with mock.patch("scripts.officecli.runtime_resolver.download_asset", side_effect=fake_download):
                with mock.patch(
                    "scripts.officecli.runtime_resolver.Path.replace",
                    side_effect=FakeWindowsFileInUseError(13, "文件正由另一进程使用", str(target)),
                ):
                    result = materialize_asset(fake_lock, fake_asset, target, offline=False, skip_version_check=True)

            self.assertEqual(result["status"], "cached")
            self.assertEqual(list(Path(tmp).glob(".download.*.tmp")), [])


class TestVersionAndLockGuards(unittest.TestCase):
    """测试版本精确匹配和 runtime 文件锁。"""

    def test_exact_version_output_accepts_only_locked_forms(self):
        """OFFICECLI-T012：版本输出只能是固定版本的精确形式。"""
        self.assertTrue(is_exact_version_output("1.0.113", "1.0.113"))
        self.assertTrue(is_exact_version_output("OfficeCLI 1.0.113", "1.0.113"))
        self.assertTrue(is_exact_version_output("v1.0.113", "1.0.113"))
        self.assertFalse(is_exact_version_output("1.0.113-dev", "1.0.113"))
        self.assertFalse(is_exact_version_output("OfficeCLI 1.0.113 latest", "1.0.113"))
        self.assertFalse(is_exact_version_output("1.0.113\nupdate available", "1.0.113"))

    def test_cli_does_not_expose_skip_version_check(self):
        """生产 CLI 不允许绕过版本校验。"""
        with self.assertRaises(SystemExit):
            parse_args(
                [
                    "ensure",
                    "--lock",
                    str(LOCK_PATH),
                    "--skip-version-check",
                ]
            )

    def test_runtime_file_lock_is_exclusive_and_released(self):
        """并发下载必须使用 runtime_id 文件锁且正常释放。"""
        with tempfile.TemporaryDirectory() as tmp:
            lock_dir = Path(tmp)
            with runtime_file_lock(lock_dir, "win-x64") as lock_path:
                self.assertTrue(lock_path.exists())
                payload = json.loads(lock_path.read_text(encoding="utf-8"))
                self.assertEqual(payload["pid"], os.getpid())
                self.assertEqual(lock_path.name, "win-x64.lock")

            self.assertFalse((lock_dir / "win-x64.lock").exists())

    def test_runtime_file_lock_removes_stale_lock_with_audit(self):
        """超过阈值且进程不存在的旧锁必须清理并写审计。"""
        with tempfile.TemporaryDirectory() as tmp:
            lock_dir = Path(tmp)
            lock_path = lock_dir / "win-x64.lock"
            lock_path.write_text(
                json.dumps({"pid": 999999999, "host": "old", "started_at": "2026-06-16T00:00:00Z"}),
                encoding="utf-8",
            )
            old = 1
            os.utime(lock_path, (old, old))

            with runtime_file_lock(lock_dir, "win-x64", wait_seconds=1):
                pass

            audit = (lock_dir / "runtime-lock.audit.jsonl").read_text(encoding="utf-8")
            self.assertIn("remove_stale_lock", audit)
            self.assertIn("win-x64", audit)


if __name__ == "__main__":
    unittest.main(verbosity=2)
