# OfficeVal Codex Runner

Run the [OmegaUse-OfficeVal](https://huggingface.co/datasets/baidu-frontier-research/OmegaUse-OfficeVal)
benchmark against any model served over the OpenAI Responses API, with a Codex agent and
LibreOffice in a throwaway Docker container.

- **One container per run.** Every task/repeat gets a fresh container, read-only rootfs,
  dropped capabilities, and CPU/memory/PID limits.
- **Bring your own model.** Point `.models.yaml` at your endpoint; one entry per model
  holds both the call config and the token prices.
- **Auditable output.** Each run archives the agent's final state, token usage, USD cost,
  redacted events, and the Office files it produced.
- **Resumable suites.** `--resume` reuses successful results and reruns the rest.

## Requirements

- Docker on `linux/amd64` (the image installs the x86-64 LibreOffice build)
- Python 3.11+ on the host
- An endpoint that speaks the OpenAI Responses API, and its API key

## Quickstart

Run every command from this `officeval_runner/` directory.

```bash
# 1. Host dependencies
python -m pip install -r requirements-host.txt

# 2. Worker image: Codex, LibreOffice, and Office/PDF/OCR Python libraries.
#    Build it in the Docker context the runner uses (--docker-context).
docker --context default build --pull -t officeval-codex:0.4.0 .

# 3. Dataset: a pinned Hugging Face revision, cached outside the repo
python prepare_dataset.py

# 4. Credentials: fill in base_url, api_key, and pricing
cp .models.example.yaml .models.yaml

# 5. Run one task
python run_benchmark.py --task 001 --model deepseek/deepseek-v4-pro-0813
```

## Configuration

`.models.yaml` is git-ignored and holds one entry per model. The entry's key is the model
ID sent to the provider and the value of `--model`. `.models.example.yaml` is an annotated
template listing every field.

- `base_url` and `api_key` are required.
- `pricing` is required, in USD per million tokens, and produces
  `calculated_token_cost_usd` in `result.json`. Omit `cache_read_per_million_usd` if the
  provider has no cache discount.
- `input_modalities: [text]` is for models without image input. Codex otherwise treats
  any model it does not know as vision-capable and sends it the images from its
  `view_image` tool; with this set, `view_image` is refused inside the container.
- All other fields are optional; their defaults live in `ModelSettings` in
  `src/config.py`.

Missing or unknown fields raise before any container starts. Web search is not
configurable: every run has Codex's `web_search` disabled.

## Usage

```bash
python run_benchmark.py --task 001 --model deepseek/deepseek-v4-pro-0813
python run_benchmark.py --task all --model deepseek/deepseek-v4-pro-0813 \
  --repeat 3 --concurrency 5 --tag baseline
```

One runner process drives one model. Run `python run_benchmark.py --help` for every flag
and its default.

Run directories are named `<date>-<hhmm>-<model>[-<tag>]`, with the model and tag
reduced to path-safe characters and truncated, e.g.
`runs/20260823-1536-deepseek-deepseek-v4-pro`. Same-minute collisions get a `-2`, `-3`
suffix.

### Resuming

```bash
python run_benchmark.py --resume runs/20260823-1536-deepseek-deepseek-v4-pro
python run_benchmark.py --resume runs/20260823-1536-deepseek-deepseek-v4-pro --force-task 001
python run_benchmark.py --resume runs/20260823-1536-deepseek-deepseek-v4-pro --concurrency 5
```

Resume keeps successful runs and reruns failed or missing ones from a fresh workspace,
using the credentials currently in `.models.yaml`. `--force-task` deletes a task's
results, with no backup, and reruns it.

The suite's config comes from its `suite.json`. Runtime knobs such as `--concurrency`,
`--max-steps`, `--timeout-seconds`, `--effort`, and the Docker options can be overridden
for one invocation. Options that define the suite (`--task`, `--model`, `--dataset-root`,
`--task-language`, `--repeat`, `--output-root`, `--tag`) cannot, and passing them with
`--resume` is an error.

## Output

Each `task-<id>/repeat-<index>/` contains:

| Path | Contents |
| --- | --- |
| `result.json` | Final agent state, usage, USD cost, archived file paths |
| `events.jsonl` | Redacted Codex events |
| `container.stderr.log` | Redacted container stderr |
| `workspace/` | The task workspace |
| `outputs/` | Office files the agent created or modified |

The suite root holds `suite.json` (config and resolved task IDs) and a `summary.json`
written when the suite finishes.

A run ends in one of `success`, `agent_failed`, `max_steps_exceeded`, `timeout`,
`container_oom`, or `container_failed`; a failed run never stops the rest of the suite.
Any other error is raised as a Python, Docker, or Codex SDK exception.

## Dataset

`prepare_dataset.py` downloads a pinned Hugging Face revision into the HF cache (set
`HF_HOME` to move it) and prints the snapshot directory. `run_benchmark.py` uses that
snapshot by default and fails if it was never prepared; `--dataset-root <dir>` points at
another copy.

## Sandbox

Each container mounts only its own task's `/workspace` and runs with a read-only rootfs,
`cap-drop=ALL`, `no-new-privileges`, resource limits, tmpfs scratch directories, and the
host user's UID/GID. It is removed when the run ends, whatever the outcome.

What the sandbox does **not** do:

- **Restrict the network.** Containers use Docker's default network to reach the model
  endpoint, so the agent can reach anything else that network can.
- **Hide the API key from the agent.** The key is passed over stdin and redacted from
  everything the runner writes to disk, but it is present inside the container. Use a
  key dedicated to benchmarking.

The read-only rootfs means the agent cannot install compiled packages, so the image
preinstalls common Office, PDF, and OCR libraries (see `requirements-container.txt`).
The worker code is copied into the image at build time: rebuild after editing `src/`.

## Debugging provider errors

`--debug` writes Codex's own logs to each run's `container.stderr.log`, the only place
the provider's real HTTP status code appears. Bare `--debug` logs one line per HTTP
request; pass a `RUST_LOG` filter for more:

```bash
python run_benchmark.py --task 011 --model deepseek/deepseek-v4-pro-0813 --debug
python run_benchmark.py --resume runs/<suite> --debug 'codex_http_client=debug,codex_core=debug'
```

Keep the filter narrow: `codex_core=debug` adds tens of MB per run, and `hyper=trace`
logs whole request bodies, images included.

## Development

```text
src/
├── config.py          # runner and model config
├── workspace.py       # task loading, input staging, artifact archiving
├── reporting.py       # pricing, JSON writing, suite summaries
├── runner.py          # suite scheduling
└── runtime/
    ├── docker.py      # host-side Docker invocation
    ├── worker.py      # in-container Codex worker
    └── codex_base_instructions.md  # Codex's fallback prompt, used with input_modalities
```

`src` runs straight from the repo; there is no install step. Run the unit tests with:

```bash
python -m unittest discover -s tests
```

The unit tests do not exercise Docker or a real provider. After changing the image or
the worker, run one real task end to end.
