"""
Loop Profiler — Non-invasive measurement wrapper for hotspot analysis.

Measures:
  - Wall-clock time per function call
  - Loop iteration counts (via manual counter ticks)
  - Call counts per function
  - Nested call tracking (e.g. _process_row called from find_and_map_smart_headers)

Usage:
    from core.utils.loop_profiler import loop_profiler, tick

    # Wrap a function to measure time + call count
    @loop_profiler.watch("my_function")
    def my_function(...):
        for item in items:
            tick("my_function")  # count each iteration
            ...

    # Or wrap at call site without decorating:
    loop_profiler.patch(module, "function_name")

    # After run, dump results:
    loop_profiler.report()      # logs summary
    loop_profiler.to_dict()     # returns raw data
    loop_profiler.reset()       # clear for next run
"""

import time
import logging
import functools
import threading
from typing import Dict, Any, Optional, Callable
from dataclasses import dataclass, field

logger = logging.getLogger(__name__)


@dataclass
class FunctionProfile:
    """Stats for a single profiled function."""
    name: str
    call_count: int = 0
    total_time_ms: float = 0.0
    min_time_ms: float = float('inf')
    max_time_ms: float = 0.0
    loop_ticks: int = 0           # manual iteration counter
    sub_ticks: Dict[str, int] = field(default_factory=dict)  # named sub-counters


class LoopProfiler:
    """
    Singleton profiler for measuring loop efficiency.
    
    Thread-safe. Attach to functions via decorator or monkey-patch.
    Results persist until reset().
    """

    def __init__(self):
        self._profiles: Dict[str, FunctionProfile] = {}
        self._lock = threading.Lock()
        self._active_context: Optional[str] = None  # current function being profiled
        self._patches: list = []  # track monkey-patches for cleanup

    def _get_or_create(self, name: str) -> FunctionProfile:
        if name not in self._profiles:
            self._profiles[name] = FunctionProfile(name=name)
        return self._profiles[name]

    def watch(self, name: str = None):
        """
        Decorator: wraps function to measure time + call count.
        
        @loop_profiler.watch("find_headers")
        def find_and_map_smart_headers(...):
            ...
        """
        def decorator(func: Callable) -> Callable:
            label = name or f"{func.__module__}.{func.__qualname__}"

            @functools.wraps(func)
            def wrapper(*args, **kwargs):
                with self._lock:
                    profile = self._get_or_create(label)
                    profile.call_count += 1

                prev_context = self._active_context
                self._active_context = label

                start = time.perf_counter()
                try:
                    return func(*args, **kwargs)
                finally:
                    elapsed_ms = (time.perf_counter() - start) * 1000
                    with self._lock:
                        profile.total_time_ms += elapsed_ms
                        profile.min_time_ms = min(profile.min_time_ms, elapsed_ms)
                        profile.max_time_ms = max(profile.max_time_ms, elapsed_ms)
                    self._active_context = prev_context

            return wrapper
        return decorator

    def tick(self, name: str = None, sub: str = None):
        """
        Count one loop iteration.
        
        Args:
            name: Profile name. If None, uses active context (from @watch).
            sub:  Optional sub-counter name (e.g. "alias_upper_calls", "merge_scan").
        """
        label = name or self._active_context
        if not label:
            return  # no active context, silently skip

        with self._lock:
            profile = self._get_or_create(label)
            profile.loop_ticks += 1
            if sub:
                profile.sub_ticks[sub] = profile.sub_ticks.get(sub, 0) + 1

    def patch(self, module, func_name: str, label: str = None):
        """
        Monkey-patch a function in a module to wrap it with profiling.
        Non-invasive — original function unchanged, just wrapped.
        
        Args:
            module: The module object containing the function.
            func_name: Name of the function to patch.
            label: Optional label (defaults to module.func_name).
        """
        original = getattr(module, func_name)
        patch_label = label or f"{module.__name__}.{func_name}"

        @functools.wraps(original)
        def patched(*args, **kwargs):
            with self._lock:
                profile = self._get_or_create(patch_label)
                profile.call_count += 1

            prev_context = self._active_context
            self._active_context = patch_label

            start = time.perf_counter()
            try:
                return original(*args, **kwargs)
            finally:
                elapsed_ms = (time.perf_counter() - start) * 1000
                with self._lock:
                    profile.total_time_ms += elapsed_ms
                    profile.min_time_ms = min(profile.min_time_ms, elapsed_ms)
                    profile.max_time_ms = max(profile.max_time_ms, elapsed_ms)
                self._active_context = prev_context

        setattr(module, func_name, patched)
        self._patches.append((module, func_name, original))
        logger.info(f"[LoopProfiler] Patched {patch_label}")

    def unpatch_all(self):
        """Restore all monkey-patched functions to originals."""
        for module, func_name, original in self._patches:
            setattr(module, func_name, original)
        count = len(self._patches)
        self._patches.clear()
        logger.info(f"[LoopProfiler] Restored {count} patched function(s)")

    def reset(self):
        """Clear all profiling data."""
        with self._lock:
            self._profiles.clear()
            self._active_context = None

    def to_dict(self) -> Dict[str, Any]:
        """Export all profiles as a dict (JSON-safe)."""
        with self._lock:
            result = {}
            for name, p in self._profiles.items():
                avg_ms = p.total_time_ms / p.call_count if p.call_count > 0 else 0
                result[name] = {
                    "calls": p.call_count,
                    "total_ms": round(p.total_time_ms, 3),
                    "avg_ms": round(avg_ms, 3),
                    "min_ms": round(p.min_time_ms, 3) if p.min_time_ms != float('inf') else 0,
                    "max_ms": round(p.max_time_ms, 3),
                    "loop_ticks": p.loop_ticks,
                    "ticks_per_call": round(p.loop_ticks / p.call_count, 1) if p.call_count > 0 else 0,
                    "sub_ticks": dict(p.sub_ticks) if p.sub_ticks else {}
                }
            return result

    def report(self, title: str = "Loop Profiler Report"):
        """
        Log a formatted summary and write to dedicated profiler log file.
        
        Writes to BOTH:
          1. Standard logger (goes to invoice_generator.log)
          2. Dedicated run_log/profiler_report.log (never wiped by session clear)
        """
        data = self.to_dict()
        if not data:
            logger.info(f"[{title}] No profiling data collected.")
            return

        lines = [
            f"{'='*70}",
            f"  {title}",
            f"{'='*70}",
            f"  {'Function':<40} {'Calls':>6} {'Total ms':>10} {'Avg ms':>9} {'Ticks':>8} {'Ticks/Call':>10}",
            f"  {'-'*40} {'-'*6} {'-'*10} {'-'*9} {'-'*8} {'-'*10}",
        ]

        for name, stats in sorted(data.items(), key=lambda x: x[1]['total_ms'], reverse=True):
            lines.append(
                f"  {name:<40} {stats['calls']:>6} {stats['total_ms']:>10.1f} "
                f"{stats['avg_ms']:>9.2f} {stats['loop_ticks']:>8} {stats['ticks_per_call']:>10.1f}"
            )
            if stats.get('sub_ticks'):
                for sub_name, sub_count in sorted(stats['sub_ticks'].items()):
                    lines.append(f"    └─ {sub_name}: {sub_count:,}")

        lines.append(f"{'='*70}")
        
        report_text = "\n".join(lines)
        
        # 1. Standard logger output
        logger.info(f"\n{report_text}\n")
        
        # 2. Write to dedicated profiler log (survives session log wipes)
        try:
            from core.system_config import sys_config
            from datetime import datetime
            
            profiler_log = sys_config.run_log_dir / "profiler_report.log"
            profiler_log.parent.mkdir(parents=True, exist_ok=True)
            
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open(profiler_log, "a", encoding="utf-8") as f:
                f.write(f"\n[{timestamp}]\n{report_text}\n")
        except Exception as e:
            logger.warning(f"[LoopProfiler] Could not write to profiler_report.log: {e}")


# --- Module-level singleton ---
loop_profiler = LoopProfiler()

# Convenience shortcut
def tick(name: str = None, sub: str = None):
    """Shortcut for loop_profiler.tick()"""
    loop_profiler.tick(name, sub)
