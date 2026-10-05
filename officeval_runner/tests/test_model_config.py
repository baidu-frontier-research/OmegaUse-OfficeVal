from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

import yaml

from src.config import load_model_config

_PRICING = {
    "input_per_million_usd": 1.0,
    "output_per_million_usd": 2.0,
    "cache_read_per_million_usd": 0.1,
}


class ModelConfigTests(unittest.TestCase):
    def _load(self, payload: object, model: str = "model-a"):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        path = Path(temporary.name) / ".models.yaml"
        path.write_text(yaml.safe_dump(payload), encoding="utf-8")
        return load_model_config(path, model)

    def test_loads_model_with_defaults(self) -> None:
        loaded = self._load(
            {
                "models": {
                    "model-a": {
                        "base_url": "https://example.com/v1/",
                        "api_key": "test-secret-value",
                        "pricing": _PRICING,
                    }
                }
            }
        )
        settings = loaded.settings

        self.assertEqual(settings.base_url, "https://example.com/v1/")
        self.assertIsNone(settings.context_window)
        self.assertEqual(settings.request_max_retries, 4)
        self.assertEqual(settings.stream_max_retries, 5)
        self.assertEqual(settings.stream_idle_timeout_ms, 300_000)
        self.assertIsNone(settings.input_modalities)
        self.assertNotIn("test-secret-value", repr(settings))

    def test_reads_prices_from_the_same_model_entry(self) -> None:
        loaded = self._load(
            {
                "models": {
                    "model-a": {
                        "base_url": "https://gateway.example/v1",
                        "api_key": "key-value",
                        "pricing": {
                            "input_per_million_usd": 1.74,
                            "output_per_million_usd": 3.48,
                        },
                    }
                }
            }
        )

        self.assertEqual(loaded.pricing.input_per_million_usd, 1.74)
        self.assertEqual(loaded.pricing.output_per_million_usd, 3.48)
        # Providers without a cache discount may omit the rate.
        self.assertEqual(loaded.pricing.cache_read_per_million_usd, 0.0)

    def test_accepts_explicit_context_and_retry_values(self) -> None:
        settings = self._load(
            {
                "models": {
                    "model-a": {
                        "base_url": "https://example.com/custom/v1",
                        "api_key": "real-looking-key",
                        "context_window": 123456,
                        "request_max_retries": 0,
                        "stream_max_retries": 0,
                        "stream_idle_timeout_ms": 1,
                        "input_modalities": ["text"],
                        "pricing": _PRICING,
                    }
                }
            }
        ).settings
        self.assertEqual(settings.context_window, 123456)
        self.assertEqual(settings.request_max_retries, 0)
        # A YAML list would make the frozen dataclass unhashable.
        self.assertEqual(settings.input_modalities, ("text",))

    def test_uses_native_mapping_and_constructor_errors(self) -> None:
        with self.assertRaisesRegex(KeyError, "models"):
            self._load({"modelEnv": {}})
        selected = self._load(
            {
                "extra": True,
                "models": {
                    "model-a": {
                        "base_url": "https://gateway.example/v1",
                        "api_key": "selected-key",
                        "pricing": _PRICING,
                    },
                    "model-b": {
                        "base_url": "https://gateway.example/v1",
                        "api_key": "other-key",
                        "unknown": True,
                        "pricing": _PRICING,
                    },
                },
            },
            model="model-a",
        )
        self.assertEqual(selected.settings.api_key, "selected-key")
        with self.assertRaisesRegex(TypeError, "temperature"):
            self._load(
                {
                    "models": {
                        "model-a": {
                            "base_url": "https://gateway.example/v1",
                            "api_key": "key-value",
                            "temperature": 1,
                            "pricing": _PRICING,
                        }
                    }
                }
            )
        # The benchmark always runs without web search; a config cannot turn it on.
        with self.assertRaisesRegex(TypeError, "web_search"):
            self._load(
                {
                    "models": {
                        "model-a": {
                            "base_url": "https://gateway.example/v1",
                            "api_key": "key-value",
                            "web_search": "live",
                            "pricing": _PRICING,
                        }
                    }
                }
            )
        with self.assertRaisesRegex(KeyError, "missing-model"):
            self._load(
                {
                    "models": {
                        "model-a": {
                            "base_url": "https://gateway.example/v1",
                            "api_key": "key-value",
                            "pricing": _PRICING,
                        }
                    }
                },
                model="missing-model",
            )

    def test_a_model_without_prices_fails_before_the_run_starts(self) -> None:
        with self.assertRaisesRegex(KeyError, "pricing"):
            self._load(
                {
                    "models": {
                        "model-a": {
                            "base_url": "https://gateway.example/v1",
                            "api_key": "key-value",
                        }
                    }
                }
            )

    def test_preserves_provider_values(self) -> None:
        settings = self._load(
            {
                "models": {
                    "model-a": {
                        "base_url": "http://gateway.example/v1/",
                        "api_key": "legitimate-token-containing-changeme-text",
                        "pricing": _PRICING,
                    }
                }
            }
        ).settings
        self.assertEqual(settings.base_url, "http://gateway.example/v1/")
        self.assertEqual(
            settings.api_key,
            "legitimate-token-containing-changeme-text",
        )


if __name__ == "__main__":
    unittest.main()
