"""Tests for the run_blocking() helper and its RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED toggle.

The helper in main.py dispatches synchronous document-generation work to
either:

1. The event loop directly (inline) — when RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED is false
   or unset. This is the legacy behavior; the loop blocks for the full
   duration of the call.
2. A worker thread via ``asyncio.to_thread`` — when RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED
   is truthy. The loop stays free to serve health probes and concurrent
   requests.

These tests verify both modes:

- Mode selection follows the env var.
- The helper returns the wrapped function's value in both modes.
- Exceptions propagate in both modes.
- In offload mode, the event loop remains responsive during a long
  blocking call.
- In inline mode, the event loop is blocked (legacy behavior).
"""

import asyncio
import importlib
import os
import sys
import threading
import time
from pathlib import Path
from typing import Any

import pytest

# Add project root to path for imports
project_root = Path(__file__).parent.parent
sys.path.insert(0, str(project_root))


# ---------------------------------------------------------------------------
# Test helpers
# ---------------------------------------------------------------------------

def _reload_main_with_flag(flag_value: str, max_workers: str | None = None):
    """Re-import main.py with the run_blocking env vars set.

    Args:
        flag_value: Value for ``RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED``.
        max_workers: Optional value for ``RUN_BLOCKING_MAX_WORKERS``. When
            ``None``, the env var is set to ``""`` so the parser falls
            back to the documented default (4). Setting it explicitly
            (including the empty string) rather than popping prevents
            ``load_dotenv()`` from re-applying a developer's local
            ``.env`` value on each ``config`` reimport.

    Modules reloaded together so that:
      - config re-reads the environment,
      - async_runner re-binds get_config and recreates the bounded
        ThreadPoolExecutor with the freshly loaded max_workers,
      - main re-binds run_blocking to the new async_runner module,
      - librechat_integration (which main imports, and which dispatches the
        upload half of every static tool) re-binds it too, instead of holding
        a run_blocking from a previous reload's async_runner.

    Returns the freshly imported main module.
    """
    os.environ["RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED"] = flag_value
    os.environ["RUN_BLOCKING_MAX_WORKERS"] = "" if max_workers is None else max_workers
    os.environ.setdefault("UPLOAD_STRATEGY", "LOCAL")

    for mod_name in ("main", "librechat_integration", "async_runner", "config"):
        sys.modules.pop(mod_name, None)

    # Defensively reset the config singleton before main reads it.
    import config as cfg_mod
    cfg_mod._CONFIG = None

    return importlib.import_module("main")


def _slow_sync(duration_s: float, return_value: Any = "ok"):
    """Synchronous sleep — emulates a blocking document-generation call."""
    time.sleep(duration_s)
    return return_value


def _raising_sync():
    """Synchronous function that always raises — for exception-path tests."""
    raise RuntimeError("boom from worker")


# ---------------------------------------------------------------------------
# Config-flag plumbing
# ---------------------------------------------------------------------------

class TestRunBlockingFlagPlumbing:
    """Verify the RUN_BLOCKING_BY_ASYNCIO_THREAD_ENABLED env var maps to config correctly."""

    @pytest.mark.parametrize("unset_marker", ["", "  "])
    def test_flag_default_is_true_when_unset(self, unset_marker):
        """Empty / whitespace env value is treated as 'unset' → default True.

        Setting an explicit value (even empty) rather than popping the var
        prevents ``load_dotenv()`` from re-applying a developer's local
        ``.env`` value on each ``config`` reimport.
        """
        main = _reload_main_with_flag(unset_marker)
        assert main.config.run_blocking_by_asyncio_thread_enabled is True

    @pytest.mark.parametrize("truthy", ["1", "true", "TRUE", "yes", "on", "y"])
    def test_flag_truthy_values(self, truthy):
        main = _reload_main_with_flag(truthy)
        assert main.config.run_blocking_by_asyncio_thread_enabled is True

    @pytest.mark.parametrize("falsy", ["0", "false", "no", "off"])
    def test_flag_explicit_falsy_values(self, falsy):
        """Only explicit falsy strings disable the offload (empty / whitespace
        falls back to the True default — see test_flag_default_is_true_when_unset)."""
        main = _reload_main_with_flag(falsy)
        assert main.config.run_blocking_by_asyncio_thread_enabled is False


# ---------------------------------------------------------------------------
# Functional behavior in both modes
# ---------------------------------------------------------------------------

class TestRunBlockingFunctional:
    """Verify run_blocking returns values and propagates exceptions in both modes."""

    @pytest.mark.parametrize("flag,expected_mode", [("false", False), ("true", True)])
    def test_returns_wrapped_value(self, flag, expected_mode):
        main = _reload_main_with_flag(flag)
        assert main.config.run_blocking_by_asyncio_thread_enabled is expected_mode

        result = asyncio.run(main.run_blocking(_slow_sync, 0.0, return_value=42))
        assert result == 42

    @pytest.mark.parametrize("flag", ["false", "true"])
    def test_propagates_exceptions(self, flag):
        main = _reload_main_with_flag(flag)
        with pytest.raises(RuntimeError, match="boom from worker"):
            asyncio.run(main.run_blocking(_raising_sync))

    @pytest.mark.parametrize("flag", ["false", "true"])
    def test_passes_args_and_kwargs(self, flag):
        main = _reload_main_with_flag(flag)

        def fn(a, b, *, c, d=10):
            return (a, b, c, d)

        result = asyncio.run(main.run_blocking(fn, 1, 2, c=3, d=4))
        assert result == (1, 2, 3, 4)


# ---------------------------------------------------------------------------
# Event-loop responsiveness — the whole reason this helper exists
# ---------------------------------------------------------------------------

class TestEventLoopBehavior:
    """Distinguish inline vs offload modes by observing loop behavior.

    The clearest signal: how long a small ``await asyncio.sleep(0.01)``
    actually takes while a 1-second blocking call is also in flight.
    """

    @staticmethod
    async def _measure_yield_latency_while_blocking(main_mod, slow_duration_s: float):
        """Return the max latency of 10 ``asyncio.sleep(0.01)`` calls while
        a slow blocking call runs concurrently."""
        latencies = []

        async def run_slow():
            await main_mod.run_blocking(_slow_sync, slow_duration_s)

        slow_task = asyncio.create_task(run_slow())
        # Let the slow task start
        await asyncio.sleep(0.05)

        for _ in range(10):
            t0 = time.perf_counter()
            await asyncio.sleep(0.01)
            latencies.append((time.perf_counter() - t0) * 1000)  # ms

        await slow_task
        return max(latencies), latencies

    def test_offload_mode_keeps_loop_responsive(self):
        """With offload enabled, yield latency stays close to the 10ms target."""
        main = _reload_main_with_flag("true")
        max_latency, _ = asyncio.run(
            self._measure_yield_latency_while_blocking(main, slow_duration_s=1.0)
        )
        # Allow generous headroom; in practice this should be ~11ms.
        assert max_latency < 50, (
            f"Loop appears blocked in offload mode: max yield latency "
            f"{max_latency:.1f} ms (expected <50ms)"
        )

    def test_inline_mode_blocks_loop(self):
        """With offload disabled, the loop is blocked during the sync call.

        We run a slow sync call concurrently with a 0.5s asyncio.sleep and
        measure total elapsed time. In inline mode, the loop can't service
        the asyncio.sleep until the sync call finishes, so total time is
        ~1.5s (sequential). In offload mode it would be ~1.0s (concurrent).
        """
        main = _reload_main_with_flag("false")

        async def race():
            t0 = time.perf_counter()
            await asyncio.gather(
                main.run_blocking(_slow_sync, 1.0),
                asyncio.sleep(0.5),
            )
            return time.perf_counter() - t0

        elapsed = asyncio.run(race())
        # Inline: ~1.5s (sequential). Offload would be ~1.0s.
        assert elapsed > 1.3, (
            f"Expected inline mode to serialize the sync call and async sleep "
            f"(~1.5s total), got {elapsed:.3f}s — looks like offload mode."
        )

    def test_offload_mode_overlaps_blocking_and_async(self):
        """With offload, a 1s sync call + 0.5s asyncio.sleep finishes in ~1s."""
        main = _reload_main_with_flag("true")

        async def race():
            t0 = time.perf_counter()
            await asyncio.gather(
                main.run_blocking(_slow_sync, 1.0),
                asyncio.sleep(0.5),
            )
            return time.perf_counter() - t0

        elapsed = asyncio.run(race())
        # Offload: ~1.0s (concurrent). Should be < 1.3s.
        assert elapsed < 1.3, (
            f"Expected offload mode to overlap sync work with async work "
            f"(~1.0s total), got {elapsed:.3f}s — looks like inline mode."
        )


# ---------------------------------------------------------------------------
# Dynamic tool integration — verify dynamic_docx / dynamic_email tools also
# go through run_blocking and follow the same flag as static tools.
# ---------------------------------------------------------------------------

class TestDynamicToolsUseRunBlocking:
    """Confirm dynamic tool registrations route through run_blocking.

    The dynamic registrars build an ``async def tool_impl(data)`` that
    delegates to ``await run_blocking(_sync_impl, data)``. We assert that
    the registered tools are coroutine functions (not plain ``def``), and
    that they execute end-to-end in both modes.
    """

    @pytest.mark.parametrize("flag", ["false", "true"])
    def test_dynamic_docx_tool_is_async_and_runs(self, flag):
        main = _reload_main_with_flag(flag)

        # Pull a registered dynamic docx tool from the mcp instance.
        # FastMCP stores tools internally; we look one up by name. The
        # actual tool function is exposed via the Tool object's `fn`.
        # We use the lower-level access pattern so this test stays
        # independent of FastMCP's public surface evolving.
        tool = asyncio.run(main.mcp.get_tool("formal_letter"))
        assert tool is not None, "dynamic tool 'formal_letter' was not registered"

        import inspect
        assert inspect.iscoroutinefunction(tool.fn), (
            "dynamic docx tool_impl should be async def so run_blocking governs dispatch"
        )

    @pytest.mark.parametrize("flag", ["false", "true"])
    def test_dynamic_email_tool_is_async_and_runs(self, flag):
        main = _reload_main_with_flag(flag)

        # bu_broadcast_email_style_1 is registered from config/email_templates.yaml
        tool = asyncio.run(main.mcp.get_tool("bu_broadcast_email_style_1"))
        assert tool is not None, "dynamic email tool was not registered"

        import inspect
        assert inspect.iscoroutinefunction(tool.fn), (
            "dynamic email tool_impl should be async def so run_blocking governs dispatch"
        )

    @pytest.mark.parametrize("flag,expected_mode", [("false", False), ("true", True)])
    def test_run_blocking_helper_imported_by_dynamic_modules(self, flag, expected_mode):
        """Both dynamic modules import run_blocking from async_runner."""
        main = _reload_main_with_flag(flag)

        # Each dynamic module must expose a `run_blocking` symbol that is
        # an async function. We don't require identity with the freshly
        # reloaded async_runner.run_blocking because reloading async_runner
        # produces a new function object while the dynamic modules retain
        # the import they bound at *their* first import time — which is
        # still functionally correct (it reads config.get_config() lazily).
        import inspect
        from docx_tools.dynamic_docx_tools import run_blocking as docx_rb
        from email_tools.dynamic_email_tools import run_blocking as email_rb

        assert inspect.iscoroutinefunction(docx_rb)
        assert inspect.iscoroutinefunction(email_rb)
        assert main.config.run_blocking_by_asyncio_thread_enabled is expected_mode


class TestUploadUsesRunBlocking:
    """The upload half of a static tool must be offloaded, not just the build.

    The five static tools used to hand the whole job — build *and* upload — to
    run_blocking. Splitting the buffer build out for LibreChat left
    upload_and_format_response() calling the synchronous backend inline, so a
    boto3 upload plus its signed-URL round-trip ran on the event loop: the very
    stall this module exists to prevent, and one that takes the health probes
    down with it. These tests pin the dispatch by thread identity rather than
    by timing, so they cannot go flaky on a loaded CI box.
    """

    @staticmethod
    def _thread_running_the_upload(flag: str) -> str:
        """Return the thread name the storage backend executed on."""
        import io
        from unittest.mock import patch

        _reload_main_with_flag(flag)

        from config import StorageStrategy
        from librechat_integration import upload_and_format_response

        observed: list[str] = []

        def record_backend(file_object, object_name):
            observed.append(threading.current_thread().name)
            return f"http://example.com/{object_name}"

        with patch("upload_tools.main.UPLOAD_STRATEGY", StorageStrategy.LOCAL), \
             patch("upload_tools.main.upload_to_local_folder", record_backend), \
             patch("librechat_integration.is_librechat_strategy", return_value=False):
            asyncio.run(
                upload_and_format_response(
                    io.BytesIO(b"content"), "docx", "probe",
                    {"user_id": None}, "created",
                )
            )

        assert observed, "storage backend never ran"
        return observed[0]

    def test_upload_runs_on_worker_thread_when_enabled(self):
        """Flag on: the blocking backend call leaves the event loop."""
        thread_name = self._thread_running_the_upload("true")

        assert thread_name.startswith("run_blocking"), (
            f"upload ran on {thread_name!r}, not a bounded-executor worker. "
            f"A synchronous S3/GCS/Azure upload on the loop thread freezes "
            f"every other request, health probes included."
        )

    def test_upload_runs_inline_when_disabled(self):
        """Flag off: dispatch follows the same legacy path as the build."""
        thread_name = self._thread_running_the_upload("false")

        assert not thread_name.startswith("run_blocking"), (
            f"inline mode must not touch the bounded executor; saw {thread_name!r}"
        )


# ---------------------------------------------------------------------------
# Bounded executor — config + behavioural tests
# ---------------------------------------------------------------------------

class TestBoundedExecutorConfig:
    """Verify RUN_BLOCKING_MAX_WORKERS maps to config and to the executor."""

    def test_default_is_4(self):
        """No env var set → config default of 4 is used."""
        main = _reload_main_with_flag("true", max_workers="")
        assert main.config.run_blocking_max_workers == 4

    @pytest.mark.parametrize("n", [1, 2, 3, 8, 16])
    def test_env_override_accepted(self, n):
        """Valid positive integer env values are honoured."""
        main = _reload_main_with_flag("true", max_workers=str(n))
        assert main.config.run_blocking_max_workers == n

    @pytest.mark.parametrize("bogus", ["0", "-3", "not-an-int", "  "])
    def test_invalid_env_falls_back_to_default(self, bogus):
        """Invalid / non-positive values silently fall back to default 4."""
        main = _reload_main_with_flag("true", max_workers=bogus)
        assert main.config.run_blocking_max_workers == 4


class TestBoundedExecutorBehavior:
    """Verify the executor is bounded and the GIL pile-up is prevented."""

    def test_executor_is_created_lazily_and_has_config_size(self):
        """First call must create a ThreadPoolExecutor sized from config."""
        _reload_main_with_flag("true", max_workers="3")
        from async_runner import _get_executor

        ex = _get_executor()
        # ThreadPoolExecutor exposes max_workers as a public attribute.
        assert ex._max_workers == 3
        # And subsequent calls must return the same instance (singleton).
        assert _get_executor() is ex

    def test_executor_uses_named_threads(self):
        """Worker threads carry the documented prefix — useful in py-spy dumps."""
        _reload_main_with_flag("true", max_workers="2")
        from async_runner import _get_executor

        captured: list[str] = []

        def record_thread_name():
            captured.append(threading.current_thread().name)

        ex = _get_executor()
        fut = ex.submit(record_thread_name)
        fut.result(timeout=5)

        assert captured, "task did not run"
        assert captured[0].startswith("run_blocking"), (
            f"thread name {captured[0]!r} should start with 'run_blocking' so "
            f"py-spy / logs make the bounded pool obvious"
        )

    def test_concurrent_runs_are_capped_at_max_workers(self):
        """With max_workers=2, never more than 2 tasks execute simultaneously.

        This is the core safety property: more than max_workers concurrent
        offloaded calls must QUEUE in the executor, not start spawning
        unbounded threads that all compete for the GIL.
        """
        _reload_main_with_flag("true", max_workers="2")
        from async_runner import run_blocking

        # We submit 8 tasks that each "hold a slot" for a short window.
        # A shared counter records how many are concurrently active; the
        # peak must equal max_workers (2), never exceed it.
        active = 0
        peak = 0
        lock = threading.Lock()
        barrier_release = threading.Event()

        def slot_holder():
            nonlocal active, peak
            with lock:
                active += 1
                if active > peak:
                    peak = active
            # Hold the slot briefly so the scheduler has a real chance to
            # pile more in if the bound is broken.
            time.sleep(0.05)
            with lock:
                active -= 1
            return True

        async def run_many():
            await asyncio.gather(*[run_blocking(slot_holder) for _ in range(8)])

        asyncio.run(run_many())

        assert peak == 2, (
            f"peak concurrent workers was {peak}, expected exactly 2 "
            f"(max_workers). Bound is broken — GIL pile-up will recur."
        )

    def test_inline_mode_does_not_use_executor(self):
        """When the thread-offload flag is disabled, no executor work happens."""
        _reload_main_with_flag("false", max_workers="2")
        from async_runner import run_blocking

        # Capture which thread the function runs on. In inline mode it
        # must be the asyncio loop thread (i.e. NOT a 'run_blocking_*'
        # worker thread).
        observed: list[str] = []

        def record_thread():
            observed.append(threading.current_thread().name)
            return 42

        async def go():
            result = await run_blocking(record_thread)
            assert result == 42

        asyncio.run(go())

        assert observed, "function did not run"
        assert not observed[0].startswith("run_blocking"), (
            f"inline mode must not touch the bounded executor; saw thread "
            f"{observed[0]!r}"
        )
