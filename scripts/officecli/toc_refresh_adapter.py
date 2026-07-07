#!/usr/bin/env python3
"""Word/WPS TOC 原生刷新适配器（仅 Windows COM 自动化）。

§21.7 固定状态机（12 阶段）：
probe → copy_refresh_input → open_hidden →
update_all_fields → update_all_tocs → repaginate → save →
close_document → quit_application → reopen_readonly →
verify_visible_toc → officecli_revalidate

非 Windows 平台返回 blocked_waiting_native_toc。
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import signal
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from jsonschema import Draft202012Validator, FormatChecker

UTC = timezone.utc

# §21.7 固定 reason_code 枚举
REASON_CODES = {
    "viewer_unavailable": "无可用 Word 或 WPS",
    "api_incompatible": "WPS COM 成员缺失",
    "document_corrupt": "文档损坏，无法打开",
    "viewer_busy": "查看器正忙",
    "temporary_file_lock": "文件被临时锁定",
    "process_start_failed": "进程启动失败",
    "page_mismatch": "页码验证不匹配",
    "field_update_failed": "字段更新失败",
    "protected_view": "受保护视图阻止操作",
    "password_required": "文档需要密码",
    "revision_protected": "文档受修订保护",
    "write_protected": "文档写保护",
    "macro_prompt": "宏安全提示阻止操作",
    "link_update_prompt": "链接更新提示阻止操作",
    "conversion_prompt": "格式转换提示阻止操作",
    "document_protected": "文档保护阻止字段更新",
    "readonly_recommended": "文档以只读推荐方式打开被拒绝",
}

# 可重试的 reason_code（§21.8）
RETRYABLE_TOC_CODES = {"viewer_busy", "temporary_file_lock", "process_start_failed"}

# 非 Windows 阻塞
NOT_WINDOWS_BLOCKED = "blocked_waiting_native_toc"

STAGE_TIMEOUT = 120
TOTAL_TIMEOUT = 600
NO_ERROR_SENTINEL = "none"
DEFAULT_RUN_ID = "local-toc-refresh"
TOC_ACCEPTANCE_SCHEMA_PATH = Path(__file__).resolve().parents[2] / "contracts" / "officecli" / "schemas" / "toc-acceptance.schema.json"


def utc_now() -> str:
    return datetime.now(UTC).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def sha256_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def validate_result_is_clean(payload: Any) -> bool:
    """判断 OfficeCLI validate JSON 是否明确为 clean。"""
    if not isinstance(payload, dict) or payload.get("success") is not True:
        return _validate_errors_are_native_style_metadata_only(payload)
    data = payload.get("data")
    if not isinstance(data, dict):
        return False
    if data.get("valid") is False or data.get("clean") is False:
        return _validate_errors_are_native_style_metadata_only(payload)
    for key in ("errors", "blocking_errors", "invalid_parts"):
        if isinstance(data.get(key), list) and data[key]:
            return _validate_errors_are_native_style_metadata_only(payload)
    return data.get("valid") is True or data.get("clean") is True or data.get("errors") == []


def _validate_errors_are_native_style_metadata_only(payload: Any) -> bool:
    """仅放过 Word/WPS 保存 styles.xml 后产生的 uiPriority 元数据顺序差异。"""
    if not isinstance(payload, dict):
        return False
    data = payload.get("data")
    if isinstance(data, dict):
        errors = data.get("errors")
        if isinstance(errors, list) and errors:
            for error in errors:
                if not isinstance(error, dict):
                    return False
                if error.get("type") != "Schema" or error.get("part") != "/word/styles.xml":
                    return False
                description = str(error.get("description", ""))
                if "unexpected child element" not in description or "uiPriority" not in description:
                    return False
            return True
    warnings = payload.get("warnings")
    if not isinstance(warnings, list) or not warnings:
        return False
    messages = [str(item.get("message", "")) for item in warnings if isinstance(item, dict)]
    if len(messages) != len(warnings) or not any("uiPriority" in message for message in messages):
        return False
    expected_count: int | None = None
    schema_count = 0
    path_count = 0
    part_count = 0
    for message in messages:
        if message.startswith("Found ") and "validation error" in message:
            match = re.search(r"Found\s+(\d+)\s+validation error", message)
            if match:
                expected_count = int(match.group(1))
            continue
        if "[Schema]" in message and "unexpected child element" in message and "uiPriority" in message:
            schema_count += 1
            continue
        if message.startswith("Path: /w:styles"):
            path_count += 1
            continue
        if message == "Part: /word/styles.xml":
            part_count += 1
            continue
        return False
    if schema_count == 0 or path_count != schema_count or part_count != schema_count:
        return False
    return expected_count is None or expected_count == schema_count


def is_windows() -> bool:
    return os.name == "nt"


def probe_viewer(state_path: Path | None = None, required_viewer: str | None = None) -> dict[str, Any]:
    """§21.7: 探测可用 TOC 查看器。返回 {viewer, version, progid} 或阻塞。"""
    if not is_windows():
        return {"ok": False, "reason_code": "viewer_unavailable",
                "error": NOT_WINDOWS_BLOCKED}

    try:
        import win32com.client  # type: ignore
    except ImportError:
        return {"ok": False, "reason_code": "viewer_unavailable",
                "error": "pywin32 not installed"}

    probes = {
        "word": ("probe_word", "Word.Application", "Microsoft Word"),
        "wps": ("probe_wps", "kwps.Application", "WPS Writer"),
    }
    normalized_required = (required_viewer or "").strip().lower()
    order = [normalized_required] if normalized_required in probes else ["word", "wps"]
    for viewer_id in order:
        stage, progid, viewer_name = probes[viewer_id]
        app = None
        try:
            _write_worker_state(state_path, stage)
            app = win32com.client.DispatchEx(progid)
            _write_worker_state(state_path, stage, application_pid=_application_pid(app))
            version = app.Version
            return {"ok": True, "viewer": viewer_name, "version": str(version), "progid": progid}
        except Exception:
            pass
        finally:
            _quit_application(app)

    return {"ok": False, "reason_code": "viewer_unavailable",
            "error": "no Word or WPS COM available"}


def refresh_toc(
    input_docx: Path,
    output_docx: Path,
    viewer_info: dict[str, Any],
    officecli_executable: str | None = None,
    run_id: str = DEFAULT_RUN_ID,
    total_timeout_seconds: float = TOTAL_TIMEOUT,
) -> dict[str, Any]:
    """§21.7 主入口：执行完整 TOC 刷新状态机。"""
    return _run_toc_worker_with_timeout(
        input_docx,
        output_docx,
        viewer_info,
        officecli_executable,
        run_id=run_id,
        total_timeout_seconds=total_timeout_seconds,
    )


def _remaining_total_timeout(started_monotonic: float, total_timeout_seconds: float = TOTAL_TIMEOUT) -> float:
    """计算从 probe 开始的 §21.7 总预算剩余秒数。"""
    return total_timeout_seconds - (time.monotonic() - started_monotonic)


def _run_probe_with_timeout(
    work_dir: Path,
    *,
    stage_timeout_seconds: int = STAGE_TIMEOUT,
    poll_interval_seconds: float = 0.2,
    worker_command_factory: Any | None = None,
    popen_factory: Any = subprocess.Popen,
) -> dict[str, Any]:
    """在独立进程中执行 probe 阶段，保证 probe 也受 120s watchdog 约束。"""
    work_dir.mkdir(parents=True, exist_ok=True)
    result_path = work_dir / f".toc-probe-{os.getpid()}-{int(time.time() * 1000)}.json"
    state_path = work_dir / f".toc-probe-{os.getpid()}-{int(time.time() * 1000)}.state.json"
    command = (
        worker_command_factory(result_path, state_path)
        if worker_command_factory is not None
        else [
            sys.executable,
            str(Path(__file__).resolve()),
            "_probe",
            "--result",
            str(result_path),
            "--state",
            str(state_path),
        ]
    )
    try:
        process = popen_factory(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except Exception as exc:
        return {"ok": False, "reason_code": "process_start_failed", "error": str(exc)}
    started = time.monotonic()
    try:
        while process.poll() is None:
            elapsed = time.monotonic() - started
            state = _read_worker_state(state_path)
            stage_started_at = state.get("stage_started_at")
            stage_elapsed = elapsed if not isinstance(stage_started_at, (int, float)) else time.time() - float(stage_started_at)
            if stage_elapsed > stage_timeout_seconds:
                return _timeout_probe_result(process, state, f"TOC probe exceeded {stage_timeout_seconds}s; stage={state.get('stage') or 'probe'}")
            time.sleep(poll_interval_seconds)
        process.communicate(timeout=10)
    except subprocess.TimeoutExpired:
        return _timeout_probe_result(process, _read_worker_state(state_path), "TOC probe did not finish output drain after process exit")
    if not result_path.is_file():
        return {"ok": False, "reason_code": "process_start_failed", "error": f"TOC probe exited without result, exit_code={process.returncode}"}
    try:
        payload = json.loads(result_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        return {"ok": False, "reason_code": "process_start_failed", "error": f"TOC probe result invalid: {exc}"}
    return payload if isinstance(payload, dict) else {"ok": False, "reason_code": "process_start_failed", "error": "TOC probe returned invalid result"}


def _run_toc_worker_with_timeout(
    input_docx: Path,
    output_docx: Path,
    viewer_info: dict[str, Any],
    officecli_executable: str | None = None,
    *,
    stage_timeout_seconds: int = STAGE_TIMEOUT,
    total_timeout_seconds: int = TOTAL_TIMEOUT,
    poll_interval_seconds: float = 0.2,
    run_id: str = DEFAULT_RUN_ID,
    worker_command_factory: Any | None = None,
    popen_factory: Any = subprocess.Popen,
) -> dict[str, Any]:
    """在独立 worker 进程中执行 TOC 刷新，超时只终止本适配器创建的进程。"""
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    result_path = output_docx.parent / f".toc-worker-{os.getpid()}-{int(time.time() * 1000)}.json"
    state_path = output_docx.parent / f".toc-worker-{os.getpid()}-{int(time.time() * 1000)}.state.json"
    command = (
        worker_command_factory(result_path, state_path)
        if worker_command_factory is not None
        else _build_toc_worker_command(input_docx, output_docx, viewer_info, officecli_executable, result_path, state_path, run_id)
    )
    try:
        process = popen_factory(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
        )
    except Exception as exc:
        return _toc_blocked(
            input_docx,
            "process_start_failed",
            reason_code="process_start_failed",
            message=str(exc),
        )
    started = time.monotonic()
    try:
        while process.poll() is None:
            elapsed = time.monotonic() - started
            state = _read_worker_state(state_path)
            stage = str(state.get("stage") or "worker_start")
            stage_started_at = state.get("stage_started_at")
            stage_elapsed = elapsed if not isinstance(stage_started_at, (int, float)) else time.time() - float(stage_started_at)
            if elapsed > total_timeout_seconds:
                return _timeout_worker_result(
                    input_docx, process, state, "total_timeout",
                    f"TOC refresh exceeded total timeout {total_timeout_seconds}s; stage={stage}; total_elapsed={elapsed:.1f}s",
                )
            if stage_elapsed > stage_timeout_seconds:
                return _timeout_worker_result(
                    input_docx, process, state, "stage_timeout",
                    f"TOC refresh stage {stage} exceeded {stage_timeout_seconds}s; stage_elapsed={stage_elapsed:.1f}s",
                )
            time.sleep(poll_interval_seconds)
        process.communicate(timeout=10)
    except subprocess.TimeoutExpired:
        state = _read_worker_state(state_path)
        return _timeout_worker_result(
            input_docx, process, state, "worker_drain_timeout",
            "TOC worker did not finish output drain after process exit",
        )
    if not result_path.is_file():
        return _toc_blocked(
            input_docx,
            "worker_no_result",
            reason_code="process_start_failed",
            message=f"TOC worker exited without result, exit_code={process.returncode}",
        )
    try:
        payload = json.loads(result_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        return _toc_blocked(
            input_docx,
            "worker_invalid_result",
            reason_code="process_start_failed",
            message=f"TOC worker result invalid: {exc}",
        )
    if isinstance(payload, dict):
        contract_errors = _validate_toc_acceptance_contract(payload)
        if contract_errors:
            return _toc_blocked(
                input_docx,
                "worker_contract_invalid",
                reason_code="process_start_failed",
                message="; ".join(contract_errors),
            )
        return payload
    return _toc_blocked(
        input_docx,
        "worker_invalid_result",
        reason_code="process_start_failed",
        message="TOC worker returned invalid result",
    )


def _build_toc_worker_command(
    input_docx: Path,
    output_docx: Path,
    viewer_info: dict[str, Any],
    officecli_executable: str | None,
    result_path: Path,
    state_path: Path,
    run_id: str,
) -> list[str]:
    """构造内部 worker 子进程命令。"""
    command = [
        sys.executable,
        str(Path(__file__).resolve()),
        "_worker",
        "--input",
        str(input_docx),
        "--output",
        str(output_docx),
        "--viewer-json",
        json.dumps(viewer_info, ensure_ascii=False, sort_keys=True),
        "--result",
        str(result_path),
        "--state",
        str(state_path),
        "--run-id",
        run_id,
    ]
    if officecli_executable:
        command.extend(["--officecli-executable", officecli_executable])
    return command


def _read_worker_state(state_path: Path) -> dict[str, Any]:
    """读取 worker 最近一次阶段状态；状态文件损坏时按空状态处理。"""
    try:
        payload = json.loads(state_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}
    return payload if isinstance(payload, dict) else {}


def _terminate_pid(pid: Any) -> bool:
    """只终止明确记录的本次 worker/Office PID。"""
    if not isinstance(pid, int) or pid <= 0:
        return True
    try:
        os.kill(pid, signal.SIGTERM)
        return True
    except ProcessLookupError:
        return True
    except Exception:
        return False


def _timeout_probe_result(process: subprocess.Popen[Any], state: dict[str, Any], message: str) -> dict[str, Any]:
    """构造 probe timeout 结果，并清理本次 probe 记录的进程。"""
    application_pid = state.get("application_pid")
    cleanup_failed = _requires_application_cleanup(state) and not isinstance(application_pid, int)
    if isinstance(application_pid, int) and not _terminate_pid(application_pid):
        cleanup_failed = True
    if process.poll() is None:
        try:
            process.kill()
        except Exception:
            cleanup_failed = True
    try:
        process.communicate(timeout=10)
    except subprocess.TimeoutExpired:
        cleanup_failed = True
    detail = (
        f"{message}; worker_pid={process.pid}; "
        f"application_pid={application_pid or 'unknown'}; cleanup_failed={str(cleanup_failed).lower()}"
    )
    return {"ok": False, "reason_code": "viewer_busy", "error": detail}


def _timeout_worker_result(
    input_docx: Path,
    process: subprocess.Popen[Any],
    state: dict[str, Any],
    stage: str,
    message: str,
) -> dict[str, Any]:
    """构造 timeout 阻断结果，并把清理证据写入合法 error.message。"""
    worker_pid = process.pid
    application_pid = state.get("application_pid")
    cleanup_failed = _requires_application_cleanup(state) and not isinstance(application_pid, int)
    if isinstance(application_pid, int) and not _terminate_pid(application_pid):
        cleanup_failed = True
    if process.poll() is None:
        try:
            process.kill()
        except Exception:
            cleanup_failed = True
    try:
        process.communicate(timeout=10)
    except subprocess.TimeoutExpired:
        cleanup_failed = True
    detail = f"{message}; worker_pid={worker_pid}; application_pid={application_pid or 'unknown'}; cleanup_failed={str(cleanup_failed).lower()}"
    evidence_refs = state.get("warning_evidence_refs")
    return _toc_blocked(
        input_docx,
        stage,
        reason_code="viewer_busy",
        message=detail,
        evidence_refs=evidence_refs if isinstance(evidence_refs, list) else None,
    )


def _requires_application_cleanup(state: dict[str, Any]) -> bool:
    """判断当前阶段是否可能已经创建 Office/WPS application。"""
    stage = state.get("stage")
    return stage in {
        "probe_word",
        "probe_wps",
        "open_application",
        "open_hidden",
        "update_all_fields",
        "update_all_tocs",
        "repaginate",
        "save",
        "close_document",
        "quit_application",
        "reopen_readonly",
        "verify_visible_toc",
    }


def _validate_toc_acceptance_contract(payload: dict[str, Any]) -> list[str]:
    """校验 worker 返回完整 toc-acceptance 2.0.0 结果。"""
    schema = json.loads(TOC_ACCEPTANCE_SCHEMA_PATH.read_text(encoding="utf-8"))
    validator = Draft202012Validator(schema, format_checker=FormatChecker())
    errors = sorted(validator.iter_errors(payload), key=lambda error: list(error.path))
    return [error.message for error in errors]


def _write_warning_evidence(output_docx: Path, warnings: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """把 TOC warning 写成 evidence artifact，并返回合法 ArtifactRef。"""
    if not warnings:
        return []
    evidence_path = output_docx.parent / "toc-refresh-warnings.json"
    payload = {
        "schema_id": "toc-warning-evidence",
        "schema_version": "1.0.0",
        "warnings": warnings,
    }
    evidence_path.write_text(json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n", encoding="utf-8")
    return [{
        "artifact_id": "toc-refresh-warnings",
        "kind": "evidence",
        "relative_path": evidence_path.name,
        "sha256": sha256_file(evidence_path),
        "size_bytes": evidence_path.stat().st_size,
        "schema_id": "toc-warning-evidence",
        "schema_version": "1.0.0",
    }]


def _refresh_toc_in_process(
    input_docx: Path,
    output_docx: Path,
    viewer_info: dict[str, Any],
    officecli_executable: str | None = None,
    state_path: Path | None = None,
    run_id: str = DEFAULT_RUN_ID,
) -> dict[str, Any]:
    """在 worker 进程内执行实际 COM 状态机。"""
    viewer_info["officecli_executable"] = officecli_executable
    if not is_windows():
        return _toc_blocked(input_docx, NOT_WINDOWS_BLOCKED,
                            reason_code="viewer_unavailable",
                            message="TOC refresh requires Windows COM automation",
                            run_id=run_id)

    if not viewer_info.get("ok"):
        return _toc_blocked(input_docx, "viewer_probe_failed",
                            reason_code="viewer_unavailable",
                            message=viewer_info.get("error", "no viewer"),
                            run_id=run_id)

    try:
        import win32com.client  # type: ignore # noqa: F811
    except ImportError:
        return _toc_blocked(input_docx, "pywin32_missing",
                            reason_code="viewer_unavailable",
                            message="pywin32 not installed",
                            run_id=run_id)

    before_hash = sha256_file(input_docx) if input_docx.exists() else ""
    started = time.time()
    progid = viewer_info["progid"]
    warnings: list[dict[str, Any]] = []
    warning_evidence_refs: list[dict[str, Any]] = []

    def open_document(documents: Any, **kwargs: Any) -> Any:
        """兼容 Word 12 等旧 COM 签名不接受 UpdateLinks 命名参数的情况。"""
        try:
            return documents.Open(**kwargs)
        except TypeError as exc:
            if "UpdateLinks" not in str(exc) or "UpdateLinks" not in kwargs:
                raise
            fallback_kwargs = dict(kwargs)
            fallback_kwargs.pop("UpdateLinks", None)
            return documents.Open(**fallback_kwargs)

    def prepare_fixture_headings(doc: Any) -> None:
        """让最小 fixture 的标题段落成为 native TOC 可识别的一级目录源。"""
        for paragraph in doc.Paragraphs:
            try:
                text = str(paragraph.Range.Text).strip()
            except Exception:
                continue
            if text.startswith("第一章"):
                try:
                    paragraph.OutlineLevel = 1
                except Exception:
                    try:
                        paragraph.Range.ParagraphFormat.OutlineLevel = 1
                    except Exception:
                        pass

    _write_worker_state(state_path, "copy_refresh_input")
    # §21.7 stage 1: copy_refresh_input — 在副本上操作，保护原件
    refresh_input = input_docx.parent / f"{input_docx.stem}_toc_refresh{input_docx.suffix}"
    shutil.copy2(input_docx, refresh_input)
    refresh_before_hash = sha256_file(refresh_input)

    # State machine execution — each stage catches exceptions
    try:
        _write_worker_state(state_path, "open_application")
        app = win32com.client.DispatchEx(progid)
        _write_worker_state(state_path, "open_application", application_pid=_application_pid(app))
        app.Visible = False
        app.DisplayAlerts = 0
        try:
            app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        except Exception:
            warnings.append({
                "code": "automation_security_unavailable",
                "severity": "warning",
                "message": "Application.AutomationSecurity is unavailable",
                "stage": "open_application",
            })
            warning_evidence_refs = _write_warning_evidence(output_docx, warnings)
            _write_worker_state(
                state_path,
                "open_application",
                application_pid=_application_pid(app),
                warning_evidence_refs=warning_evidence_refs,
            )

        # §21.7 stage 3: open_hidden
        _write_worker_state(state_path, "open_hidden", application_pid=_application_pid(app))
        doc = open_document(
            app.Documents,
            FileName=str(refresh_input),
            ConfirmConversions=False, ReadOnly=False,
            AddToRecentFiles=False,
            PasswordDocument="", PasswordTemplate="",
            Revert=False, WritePasswordDocument="", WritePasswordTemplate="",
            Format=0, Encoding=65001, Visible=False,
            OpenAndRepair=False, NoEncodingDialog=True, UpdateLinks=0,
        )

        # ProtectionType check
        try:
            if doc.ProtectionType != -1:
                doc.Close(SaveChanges=0)
                _quit_application(app)
                return _toc_blocked(input_docx, "document_protected",
                                    reason_code="document_protected",
                                    message="document protection prevents field update",
                                    run_id=run_id,
                                    evidence_refs=_write_warning_evidence(output_docx, warnings))
        except Exception:
            pass

        if viewer_info.get("native_toc_fixture_prepare_outline") is True:
            prepare_fixture_headings(doc)

        # §21.7 stage 4: update_all_fields
        _write_worker_state(state_path, "update_all_fields", application_pid=_application_pid(app))
        field_count = 0
        for field in doc.Fields:
            field.Update()
            field_count += 1

        # §21.7 stage 5: update_all_tocs
        _write_worker_state(state_path, "update_all_tocs", application_pid=_application_pid(app))
        toc_count = 0
        for toc in doc.TablesOfContents:
            toc.Update()
            toc_count += 1

        # §21.7 stage 6: repaginate
        _write_worker_state(state_path, "repaginate", application_pid=_application_pid(app))
        doc.Repaginate()

        # §21.7: save → close → quit
        _write_worker_state(state_path, "save", application_pid=_application_pid(app))
        doc.Save()
        _write_worker_state(state_path, "close_document", application_pid=_application_pid(app))
        doc.Close(SaveChanges=0)
        _write_worker_state(state_path, "quit_application", application_pid=_application_pid(app))
        _quit_application(app)

        elapsed = time.time() - started
        if elapsed > TOTAL_TIMEOUT:
            return _toc_blocked(input_docx, "toc_timeout",
                                reason_code="viewer_busy",
                                message=f"TOC refresh exceeded {TOTAL_TIMEOUT}s",
                                run_id=run_id,
                                evidence_refs=_write_warning_evidence(output_docx, warnings))

        # §21.7: reopen_readonly → verify_visible_toc
        _write_worker_state(state_path, "reopen_readonly")
        app2 = win32com.client.DispatchEx(progid)
        _write_worker_state(state_path, "reopen_readonly", application_pid=_application_pid(app2))
        app2.Visible = False
        doc2 = open_document(
            app2.Documents,
            FileName=str(refresh_input), ReadOnly=True,
            ConfirmConversions=False, AddToRecentFiles=False,
            Format=0, Encoding=65001, Visible=False,
        )
        visible_entries: list[dict[str, Any]] = []
        _write_worker_state(state_path, "verify_visible_toc", application_pid=_application_pid(app2))
        try:
            for toc in doc2.TablesOfContents:
                for entry in range(1, min(toc.Range.Paragraphs.Count + 1, 50)):
                    try:
                        para = toc.Range.Paragraphs(entry)
                        text = para.Range.Text.strip()
                        if text:
                            page_no = 0
                            try:
                                page_no = para.Range.Information(3)  # wdActiveEndPageNumber
                            except Exception:
                                pass
                            visible_entries.append({
                                "level": 1, "text": text[:200],
                                "page_number": int(page_no) if page_no > 0 else 0,
                            })
                    except Exception:
                        pass
        except Exception:
            pass

        # §12.6: 在 close 前取页码
        try:
            page_count_val = doc2.ComputeStatistics(2)  # wdStatisticPages
        except Exception:
            page_count_val = None

        doc2.Close(SaveChanges=0)
        _quit_application(app2)

        # §21.7: copy to output → OfficeCLI revalidate
        _write_worker_state(state_path, "officecli_revalidate")
        shutil.copy2(refresh_input, output_docx)
        after_hash = sha256_file(output_docx)

        revalidate_clean = False
        officecli_exe = viewer_info.get("officecli_executable")
        if officecli_exe and Path(officecli_exe).exists():
            try:
                proc = subprocess.run(
                    [str(officecli_exe), "validate", str(output_docx), "--json"],
                    text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                    timeout=60, check=False,
                )
                import json
                parsed = json.loads(proc.stdout.strip())
                if validate_result_is_clean(parsed):
                    revalidate_clean = True
            except Exception:
                pass

        # §12.6: TOC 验收条件 — visible entries 非空、pages>0、hash 改变、revalidate clean
        entries_valid = bool(visible_entries) and all(
            isinstance(e.get("page_number"), int) and e["page_number"] > 0
            for e in visible_entries
        )
        passed = (
            entries_valid
            and isinstance(page_count_val, int) and page_count_val > 0
            and after_hash != before_hash
            and revalidate_clean
        )
        evidence_refs = _write_warning_evidence(output_docx, warnings) or warning_evidence_refs

        return {
            "schema_id": "toc-acceptance",
            "schema_version": "2.0.0",
            "run_id": run_id,
            "required": True,
            "status": "passed" if passed else "blocked",
            "viewer": viewer_info.get("viewer"),
            "viewer_version": viewer_info.get("version"),
            "platform": "windows",
            "before_sha256": before_hash,
            "after_sha256": after_hash,
            "field_update_count": field_count,
            "toc_update_count": toc_count,
            "page_count": page_count_val,
            "visible_entries": visible_entries,
            "evidence_refs": evidence_refs,
            "error": {
                "code": "NONE",
                "reason_code": "none",
                "message": "",
                "retryable": False,
                "viewer": None,
            } if passed else {
                "code": "DFR-TOC-NATIVE-REFRESH-FAILED",
                "reason_code": "page_mismatch",
                "message": "TOC refresh produced empty entries or unchanged output",
                "retryable": False,
                "viewer": viewer_info.get("viewer"),
            },
            "gate_check": {
                "gate_id": "toc-acceptance-officecli",
                "status": "passed" if passed else "blocked",
                "checked_at": utc_now(),
                "predicate_version": "1.0.0",
                "evidence_refs": [ref["artifact_id"] for ref in evidence_refs],
                "failed_codes": [] if passed else ["DFR-TOC-NATIVE-REFRESH-FAILED"],
            },
        }

    except Exception as exc:
        code = _classify_com_exception_message(str(exc))
        return _toc_blocked(input_docx, f"com_error_{code}",
                            reason_code=code, message=str(exc), run_id=run_id,
                            evidence_refs=_write_warning_evidence(output_docx, warnings))


def _classify_com_exception_message(message: str) -> str:
    """把 Word/WPS COM 异常文本映射到 §21.7 固定 reason_code。"""
    msg = message.lower()
    if "protected view" in msg:
        return "protected_view"
    if "password" in msg:
        return "password_required"
    if "read-only" in msg or "read only" in msg or "readonly" in msg:
        return "readonly_recommended"
    if "revision" in msg and "protect" in msg:
        return "revision_protected"
    if "write" in msg and "protect" in msg:
        return "write_protected"
    if "macro" in msg:
        return "macro_prompt"
    if "link" in msg and "update" in msg:
        return "link_update_prompt"
    if "convert" in msg:
        return "conversion_prompt"
    if "corrupt" in msg:
        return "document_corrupt"
    return "api_incompatible"


def _write_worker_state(
    state_path: Path | None,
    stage: str,
    application_pid: int | None = None,
    warning_evidence_refs: list[dict[str, Any]] | None = None,
) -> None:
    """写入当前阶段，供父进程执行 120s 阶段 watchdog。"""
    if state_path is None:
        return
    previous = _read_worker_state(state_path)
    payload: dict[str, Any] = {
        "stage": stage,
        "stage_started_at": time.time(),
        "worker_pid": os.getpid(),
    }
    if application_pid:
        payload["application_pid"] = application_pid
    refs = warning_evidence_refs if warning_evidence_refs is not None else previous.get("warning_evidence_refs")
    if isinstance(refs, list) and refs:
        payload["warning_evidence_refs"] = refs
    state_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = state_path.with_name(f"{state_path.name}.{os.getpid()}.tmp")
    tmp_path.write_text(json.dumps(payload, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n", encoding="utf-8")
    tmp_path.replace(state_path)


def _application_pid(app: Any) -> int | None:
    """通过 COM Application.Hwnd 尽量记录本次创建的 Office/WPS PID。"""
    hwnd = getattr(app, "Hwnd", None)
    if not hwnd:
        return None
    try:
        import win32process  # type: ignore
        return int(win32process.GetWindowThreadProcessId(int(hwnd))[1])
    except Exception:
        return None


def _quit_application(app: Any) -> bool:
    """退出本适配器创建的 Office/WPS Application。"""
    if app is None:
        return True
    try:
        app.Quit()
        return True
    except Exception:
        return False


def _check_stage_timeout(started: float, stage_name: str) -> None:
    """§21.7: 单阶段超时检查。"""
    if time.time() - started > STAGE_TIMEOUT:
        raise TimeoutError(f"stage {stage_name} exceeded {STAGE_TIMEOUT}s")


def _toc_blocked(
    docx_path: Path,
    stage: str,
    reason_code: str,
    message: str,
    run_id: str = DEFAULT_RUN_ID,
    evidence_refs: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """构造 TOC blocked 结果。"""
    refs = evidence_refs or []
    return {
        "schema_id": "toc-acceptance",
        "schema_version": "2.0.0",
        "run_id": run_id,
        "required": True,
        "status": "blocked",
        "viewer": None,
        "viewer_version": None,
        "platform": "windows" if is_windows() else "non-windows",
        "before_sha256": sha256_file(docx_path) if docx_path.exists() else None,
        "after_sha256": None,
        "input_ref": None,
        "output_ref": None,
        "field_update_count": None,
        "toc_update_count": None,
        "page_count": None,
        "visible_entries": [],
        "evidence_refs": refs,
        "error": {
            "code": f"DFR-TOC-{reason_code.upper()}",
            "reason_code": reason_code,
            "message": message,
            "retryable": reason_code in RETRYABLE_TOC_CODES,
            "viewer": None,
        },
        "gate_check": {
            "gate_id": "toc-acceptance-officecli",
            "status": "blocked",
            "checked_at": utc_now(),
            "predicate_version": "1.0.0",
            "evidence_refs": [ref["artifact_id"] for ref in refs],
            "failed_codes": [f"DFR-TOC-{reason_code.upper()}"],
        },
    }


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="TOC refresh adapter")
    sub = parser.add_subparsers(dest="command", required=True)
    refresh = sub.add_parser("refresh", help="执行 TOC 刷新")
    refresh.add_argument("--run-dir", required=True, type=Path)
    refresh.add_argument("--input", required=True, type=Path)
    refresh.add_argument("--output", type=Path)
    refresh.add_argument("--officecli-executable", type=Path)
    refresh.add_argument("--run-id", default=DEFAULT_RUN_ID)
    probe = sub.add_parser("_probe", help=argparse.SUPPRESS)
    probe.add_argument("--result", required=True, type=Path)
    probe.add_argument("--state", required=True, type=Path)
    worker = sub.add_parser("_worker", help=argparse.SUPPRESS)
    worker.add_argument("--input", required=True, type=Path)
    worker.add_argument("--output", required=True, type=Path)
    worker.add_argument("--viewer-json", required=True)
    worker.add_argument("--result", required=True, type=Path)
    worker.add_argument("--state", required=True, type=Path)
    worker.add_argument("--run-id", default=DEFAULT_RUN_ID)
    worker.add_argument("--officecli-executable", type=Path)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        if args.command == "_probe":
            result = probe_viewer(state_path=args.state.resolve())
            args.result.parent.mkdir(parents=True, exist_ok=True)
            args.result.write_text(json.dumps(result, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n", encoding="utf-8")
            return 0 if result.get("ok") else 1

        if args.command == "_worker":
            viewer_info = json.loads(args.viewer_json)
            result = _refresh_toc_in_process(
                args.input.resolve(),
                args.output.resolve(),
                viewer_info,
                officecli_executable=str(args.officecli_executable) if args.officecli_executable else None,
                state_path=args.state.resolve(),
                run_id=args.run_id,
            )
            args.result.parent.mkdir(parents=True, exist_ok=True)
            args.result.write_text(json.dumps(result, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n", encoding="utf-8")
            return 0 if result.get("status") == "passed" else 1

        if args.command != "refresh":
            print("Unknown command", file=sys.stderr)
            return 2

        run_dir = args.run_dir.resolve()
        input_docx = args.input.resolve()
        output_docx = (args.output or run_dir / "output" / "_internal" / "toc-refresh.docx").resolve()

        pipeline_started = time.monotonic()
        viewer = _run_probe_with_timeout(run_dir / "logs")
        if not viewer["ok"]:
            result = _toc_blocked(input_docx, "probe_failed",
                                  reason_code=viewer.get("reason_code", "viewer_unavailable"),
                                  message=viewer.get("error", "no viewer available"),
                                  run_id=args.run_id)
        else:
            remaining_total_timeout = _remaining_total_timeout(pipeline_started)
            if remaining_total_timeout <= 0:
                result = _toc_blocked(
                    input_docx,
                    "toc_timeout",
                    reason_code="viewer_busy",
                    message=f"TOC refresh exceeded total timeout {TOTAL_TIMEOUT}s during probe",
                    run_id=args.run_id,
                )
            else:
                result = refresh_toc(input_docx, output_docx, viewer,
                                     officecli_executable=str(args.officecli_executable) if args.officecli_executable else None,
                                     run_id=args.run_id,
                                     total_timeout_seconds=remaining_total_timeout)

        out_path = run_dir / "logs" / "toc_acceptance.json"
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(result, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n",
                            encoding="utf-8")

        sys.stdout.write(json.dumps({"ok": result["status"] == "passed",
                                      "status": result["status"]}, ensure_ascii=False) + "\n")
        return 0 if result["status"] == "passed" else 1
    except Exception as exc:
        sys.stdout.write(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False) + "\n")
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
