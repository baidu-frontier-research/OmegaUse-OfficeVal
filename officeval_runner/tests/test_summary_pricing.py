from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from src.config import PricingRates
from src.reporting import aggregate_results, calculate_token_cost, write_summary


def _result(
    *,
    usage: dict[str, int] | None,
    cost: float | None,
    status: str = "success",
) -> dict[str, object]:
    return {
        "task_id": "001",
        "model": "model-a",
        "repeat_index": 1,
        "agent": {
            "status": status,
            "wall_time_ms": 100,
            "usage": usage,
        },
        "calculated_token_cost_usd": cost,
        "output_files": ["answer.docx"],
    }


class PricingAndSummaryTests(unittest.TestCase):
    def test_prices_codex_usage_without_repricing_reasoning(self) -> None:
        usage = {
            "input_tokens": 100,
            "cached_input_tokens": 40,
            "output_tokens": 50,
            "reasoning_output_tokens": 20,
            "total_tokens": 150,
        }
        rates = PricingRates(
            input_per_million_usd=10,
            output_per_million_usd=20,
            cache_read_per_million_usd=2,
        )

        cost = calculate_token_cost(rates, usage)

        self.assertAlmostEqual(cost, (60 * 10 + 40 * 2 + 50 * 20) / 1_000_000)

    def test_pricing_uses_native_missing_field_error(self) -> None:
        with self.assertRaisesRegex(KeyError, "cached_input_tokens"):
            calculate_token_cost(
                PricingRates(1, 1),
                {"input_tokens": 10},  # type: ignore[arg-type]
            )

    def test_missing_usage_has_no_cost(self) -> None:
        self.assertIsNone(calculate_token_cost(PricingRates(1, 1), None))

    def test_summary_keeps_incomplete_usage_totals_null(self) -> None:
        usage = {
            "input_tokens": 10,
            "cached_input_tokens": 2,
            "output_tokens": 4,
            "reasoning_output_tokens": 1,
            "total_tokens": 14,
        }
        summary = aggregate_results(
            [
                _result(usage=usage, cost=0.1),
                _result(usage=None, cost=None, status="agent_failed"),
            ]
        )

        self.assertEqual(summary["model"], "model-a")
        self.assertEqual(summary["status_counts"], {"agent_failed": 1, "success": 1})
        self.assertEqual(summary["usage_missing_runs"], 1)
        self.assertIsNone(summary["total_tokens"])
        self.assertIsNone(summary["calculated_token_cost_usd_sum"])

    def test_writes_summary_json(self) -> None:
        usage = {
            "input_tokens": 1,
            "cached_input_tokens": 0,
            "output_tokens": 2,
            "reasoning_output_tokens": 1,
            "total_tokens": 3,
        }
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            write_summary(root, [_result(usage=usage, cost=0.01)])
            summary = json.loads((root / "summary.json").read_text())

        self.assertEqual(summary["runs"], 1)
        self.assertEqual(summary["total_tokens"], 3)


if __name__ == "__main__":
    unittest.main()
