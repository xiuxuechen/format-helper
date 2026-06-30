"""OfficeCLI 运行时解析器测试。

覆盖 V5-T001 到 V5-T014 中无需真实下载即可验证的供应链和平台规则。
"""

from __future__ import annotations

import hashlib
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path


sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from scripts.officecli.runtime_resolver import (  # noqa: E402
    EXPECTED_RUNTIME_IDS,
    FH_OFFICECLI_CACHE_INVALID,
    FH_OFFICECLI_OFFLINE_CACHE_MISS,
    FH_OFFICECLI_PLATFORM_UNSUPPORTED,
    OfficeCliRuntimeError,
    cache_executable_path,
    detect_runtime_id,
    is_exact_version_output,
    load_lock,
    parse_args,
    runtime_file_lock,
    select_asset,
    validate_lock,
    verify_file_hash_and_size,
)


LOCK_PATH = Path(__file__).parent.parent.parent / "tools" / "officecli" / "officecli.lock.json"


class TestOfficeCliLock(unittest.TestCase):
    """测试 OfficeCLI 锁文件。"""

    def test_lock_contains_exact_eight_assets(self):
        """V5-T001：锁文件必须精确包含八个平台资产。"""
        lock = load_lock(LOCK_PATH)
        runtime_ids = {asset["runtime_id"] for asset in lock["assets"]}

        self.assertEqual(lock["officecli_version"], "1.0.113")
        self.assertEqual(lock["release_tag"], "v1.0.113")
        self.assertEqual(runtime_ids, EXPECTED_RUNTIME_IDS)
        self.assertEqual(len(lock["assets"]), 8)
        self.assertTrue(lock["auto_update_disabled"])

    def test_lock_drift_is_rejected(self):
        """V5-T006：锁文件版本漂移必须阻塞。"""
        lock = load_lock(LOCK_PATH)
        lock["officecli_version"] = "1.0.114"

        with self.assertRaises(OfficeCliRuntimeError) as ctx:
            validate_lock(lock)

        self.assertEqual(ctx.exception.code, "FH-OFFICECLI-LOCK-INVALID")

    def test_lock_asset_tuple_drift_is_rejected(self):
        """V5-T006：任一平台官方资产元组漂移必须阻塞。"""
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
        """V5-T007：Windows x64 映射到 win-x64。"""
        self.assertEqual(detect_runtime_id("Windows", "AMD64"), "win-x64")

    def test_detect_macos_arm64(self):
        """V5-T008：macOS arm64 映射到 osx-arm64。"""
        self.assertEqual(detect_runtime_id("Darwin", "arm64"), "osx-arm64")

    def test_detect_linux_glibc(self):
        """V5-T009：Linux glibc 映射到 gnu 资产。"""
        result = detect_runtime_id(
            "Linux",
            "x86_64",
            alpine_release_exists=False,
            ldd_output="ldd (GNU libc) 2.39",
        )

        self.assertEqual(result, "linux-x64-gnu")

    def test_detect_linux_musl(self):
        """V5-T010：Linux musl/Alpine 映射到 alpine 资产。"""
        result = detect_runtime_id(
            "Linux",
            "aarch64",
            alpine_release_exists=True,
            ldd_output="",
        )

        self.assertEqual(result, "linux-arm64-musl")

    def test_unsupported_arch_is_rejected(self):
        """V5-T011：未知架构不得猜测。"""
        with self.assertRaises(OfficeCliRuntimeError) as ctx:
            detect_runtime_id("Linux", "riscv64")

        self.assertEqual(ctx.exception.code, FH_OFFICECLI_PLATFORM_UNSUPPORTED)


class TestCacheValidation(unittest.TestCase):
    """测试缓存路径和文件校验。"""

    def test_cache_path_uses_locked_version_and_runtime(self):
        """V5-T002：缓存路径必须使用固定版本和 runtime_id。"""
        lock = load_lock(LOCK_PATH)
        asset = select_asset(lock, "win-x64")
        with tempfile.TemporaryDirectory() as tmp:
            path = cache_executable_path(Path(tmp), lock, asset)

        self.assertTrue(str(path).replace("\\", "/").endswith(".cache/officecli/v1.0.113/win-x64/officecli.exe"))

    def test_missing_offline_cache_is_rejected(self):
        """V5-T013：无网且缓存不存在必须阻塞。"""
        lock = load_lock(LOCK_PATH)
        asset = select_asset(lock, "win-x64")
        with tempfile.TemporaryDirectory() as tmp:
            missing = Path(tmp) / "officecli.exe"
            with self.assertRaises(OfficeCliRuntimeError) as ctx:
                verify_file_hash_and_size(missing, asset)

        self.assertEqual(ctx.exception.code, FH_OFFICECLI_OFFLINE_CACHE_MISS)

    def test_cache_size_mismatch_is_rejected(self):
        """V5-T014：缓存损坏必须阻塞。"""
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


class TestVersionAndLockGuards(unittest.TestCase):
    """测试版本精确匹配和 runtime 文件锁。"""

    def test_exact_version_output_accepts_only_locked_forms(self):
        """V5-T012：版本输出只能是固定版本的精确形式。"""
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
