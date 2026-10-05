from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml


@dataclass(frozen=True)
class PricingRates:
    input_per_million_usd: float
    output_per_million_usd: float
    cache_read_per_million_usd: float = 0.0


@dataclass(frozen=True)
class ModelSettings:
    base_url: str
    api_key: str = field(repr=False, compare=False)
    context_window: int | None = None
    request_max_retries: int = 4
    stream_max_retries: int = 5
    stream_idle_timeout_ms: int = 300_000
    input_modalities: tuple[str, ...] | None = None

    def __post_init__(self) -> None:
        # YAML hands us a list; keep the dataclass immutable and hashable.
        if self.input_modalities is not None:
            object.__setattr__(self, "input_modalities", tuple(self.input_modalities))


@dataclass(frozen=True)
class ModelConfig:
    """One model's entry: how to call the provider, and what its tokens cost."""

    settings: ModelSettings
    pricing: PricingRates


@dataclass(frozen=True)
class RunnerConfig:
    dataset_root: Path
    output_root: Path
    model: str
    task_ids: tuple[str, ...]
    models_path: Path
    task_language: str = "zh"
    repeats: int = 1
    concurrency: int = 1
    max_steps: int = 200
    timeout_seconds: float = 7200.0
    docker_context: str = "default"
    docker_image: str = "officeval-codex:0.4.0"
    container_cpus: float = 2.0
    container_memory: str = "4g"
    container_pids_limit: int = 512
    effort: str | None = None
    debug: str | None = None
    tag: str | None = None


def load_model_config(path: Path, model: str) -> ModelConfig:
    """Load one model's provider settings and token prices."""
    models = yaml.safe_load(path.resolve().read_text(encoding="utf-8"))["models"]
    entry = dict(models[model])
    pricing = PricingRates(**entry.pop("pricing"))
    return ModelConfig(settings=ModelSettings(**entry), pricing=pricing)
