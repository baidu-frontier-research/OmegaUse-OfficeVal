from __future__ import annotations

import json
import os
from collections import Counter
from pathlib import Path
from typing import Any, Iterable

from .config import PricingRates


def calculate_token_cost(
    rates: PricingRates, usage: dict[str, int] | None
) -> float | None:
    if usage is None:
        return None
    cached_input = usage["cached_input_tokens"]
    uncached_input = usage["input_tokens"] - cached_input
    return (
        uncached_input * rates.input_per_million_usd
        + usage["output_tokens"] * rates.output_per_million_usd
        + cached_input * rates.cache_read_per_million_usd
    ) / 1_000_000


def write_json(path: Path, payload: Any) -> None:
    """Atomically write a JSON document to disk."""
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.tmp")
    temporary.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    os.replace(temporary, path)


def _usage_sum(results: list[dict[str, Any]], field: str) -> int | None:
    values = [
        result["agent"]["usage"][field]
        if result["agent"]["usage"] is not None
        else None
        for result in results
    ]
    if any(value is None for value in values):
        return None
    return sum(values)


def aggregate_results(results: Iterable[dict[str, Any]]) -> dict[str, Any]:
    """Aggregate one suite's run records."""
    results = list(results)
    total = len(results)
    successful = sum(
        result["agent"]["status"] == "success" for result in results
    )
    usage_available = sum(
        result["agent"]["usage"] is not None for result in results
    )
    costs = [result["calculated_token_cost_usd"] for result in results]
    return {
        "model": results[0]["model"],
        "runs": total,
        "status_counts": dict(
            sorted(Counter(result["agent"]["status"] for result in results).items())
        ),
        "agent_successes": successful,
        "agent_success_rate": successful / total,
        "runs_with_saved_outputs": sum(bool(result["output_files"]) for result in results),
        "total_wall_time_ms": sum(
            result["agent"]["wall_time_ms"] for result in results
        ),
        "usage_available_runs": usage_available,
        "usage_missing_runs": total - usage_available,
        "total_input_tokens": _usage_sum(results, "input_tokens"),
        "total_cached_input_tokens": _usage_sum(results, "cached_input_tokens"),
        "total_output_tokens": _usage_sum(results, "output_tokens"),
        "total_reasoning_output_tokens": _usage_sum(
            results, "reasoning_output_tokens"
        ),
        "total_tokens": _usage_sum(results, "total_tokens"),
        "calculated_token_cost_usd_sum": (
            None if any(cost is None for cost in costs) else sum(costs)
        ),
    }


def write_summary(suite_dir: Path, results: list[dict[str, Any]]) -> None:
    """Write the aggregate summary for a suite."""
    write_json(suite_dir / "summary.json", aggregate_results(results))
