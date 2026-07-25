# OmegaUse-OfficeVal

**Benchmarking LLM Agents on Long-Horizon Office-Suite Tasks with Economic Grounding**

OmegaUse-OfficeVal is a benchmark for evaluating LLM agents on long-horizon,
real-world office-suite tasks spanning word-processing documents, spreadsheets,
presentations, and cross-file productivity workflows. The tasks are derived from
authentic office requests proposed by practitioners and adapted through a
privacy-preserving pipeline that removes sensitive information while preserving
the original user intent, constraints, and natural request phrasing.

A distinctive feature of the benchmark is its **economic grounding**: every task
is annotated with two complementary signals — *human labor time* and a
*task price proxy* — enabling evaluation not only of task completion, but also of
the human effort and economic value associated with each task.

The dataset comprises **100 tasks** collected from real office scenarios.

## Directory Structure

Each task is described by two parallel JSON files that share the same identifier,
`officeval_001` through `officeval_100`:

```
data/
  tasks/
    officeval_001.json      # task metadata (instruction, economic signals, input files)
    ...
    officeval_100.json
  rubrics/
    officeval_001.json      # evaluation rubric for the corresponding task
    ...
    officeval_100.json
  README.md
```

- **`tasks/`** — One file per task holding the instruction, task attributes,
  economic-grounding signals, and the list of input files.
- **`rubrics/`** — One file per task holding the evaluation rubric. A rubric file
  matches its task file by identical `id` / filename.

The input artifacts referenced by each task are hosted on the Hugging Face Hub:

```
https://huggingface.co/datasets/baidu-frontier-research/OmegaUse-OfficeVal
```

## Task Schema (`tasks/officeval_<NNN>.json`)

- **`id`** *(string)* — Task identifier, `"officeval_001"`–`"officeval_100"`.

- **`instruction`** *(string)* — The natural-language user request describing the
  task. Instructions preserve the original colloquial style and phrasing of real
  office requests; line breaks separate individual sub-requirements.

- **`operation_intent`** *(string)* — The primary operation the task requires.
  One of: `Reformat`, `Restructure`, `Annotate`, `Extract`, `Compute`,
  `Beautify`, `Other`.

- **`domain`** *(string)* — The professional domain of the task. One of:
  `Academic Papers`, `Education & Examination`, `Financial Data`,
  `Engineering & Technology`, `Administrative Affairs`, `Business Operations`,
  `Other`.

- **`human_labor_time`** *(number)* — **Unit: minutes (min).** The recorded time
  required by recruited human annotators to complete the task *without* LLM
  assistance. Each task is completed by at least two annotators under quality
  control, and the reported value is the average of the two shortest valid
  completion times.

- **`task_price_proxy`** *(number)* — **Unit: yuan (CNY).** A task-level price
  signal estimating the market price of completing the task. It reflects the
  perceived importance / willingness to pay and captures the economic value of the
  task.

- **`price_source`** *(string)* — Indicates how `task_price_proxy` was obtained:
  - `explicit_price` — a real price signal provided by practitioners
    (e.g., a task previously outsourced on a freelance platform).
  - `estimated_price` — an expert estimate produced through a consistency-based
    aggregation of independent expert annotations, used when no explicit price is
    available.

- **`origin_files`** *(array of objects)* — The input files required to complete
  the task. Each entry contains:
  - **`url`** — A direct-download link to the file on the Hugging Face Hub
    (`.../resolve/main/<NNN>/<filename>`).
  - **`dest`** — The local filename the artifact should be saved as.

## Rubric Schema (`rubrics/officeval_<NNN>.json`)

- **`id`** *(string)* — Matches the corresponding task id, e.g. `"officeval_001"`.

- **`instruction`** *(string)* — The same instruction text as the task file,
  duplicated so the rubric is self-contained.

- **`rubrics`** *(object)* — The task evaluation rubric, split into two dimensions.
  Each dimension is a list of checkable requirement strings.
  - **`dim1`** — *Usability rubric.* Basic usability requirements that the
    delivered artifact must satisfy (e.g., correct file format, the file opens
    normally, content/layout is not corrupted, the file remains editable). If any
    usability item fails, the task scores zero and the completion rubric is not
    evaluated further.
  - **`dim2`** — *Task-completion rubric.* Fine-grained, user-centered scoring
    points that measure how well the specific task requirements are met. Positive
    items (prefixed `+n`) reward correctly completed requirements; negative items
    (prefixed `-n`) penalize unintended changes or avoidable damage that increase
    the user's downstream repair cost. The leading integer denotes the discrete
    weight of the item.

### Example

`tasks/officeval_001.json`

```json
{
  "id": "officeval_001",
  "instruction": "按照下述要求修改给到你的word文档\n1、...",
  "operation_intent": "Restructure",
  "domain": "Education & Examination",
  "human_labor_time": 204,
  "task_price_proxy": 50,
  "price_source": "estimated_price",
  "origin_files": [
    {
      "url": "https://huggingface.co/datasets/baidu-frontier-research/OmegaUse-OfficeVal/resolve/main/001/01_课时学习方案_观察一位校园志愿者_第一课时.docx",
      "dest": "01_课时学习方案_观察一位校园志愿者_第一课时.docx"
    }
  ]
}
```

`rubrics/officeval_001.json`

```json
{
  "id": "officeval_001",
  "instruction": "按照下述要求修改给到你的word文档\n1、...",
  "rubrics": {
    "dim1": ["交付文件为Word格式，扩展名为.doc，文件可正常打开。", "..."],
    "dim2": ["+1：学历案内含“课时学习目标”标题且标题下无内容，...", "..."]
  }
}
```

## Dataset Statistics

- **Tasks:** 100
- **Output modalities:** Word, PowerPoint, Excel, PDF, and cross-file tasks.
- **Operation intent:** Other (43), Reformat (29), Restructure (14), Extract (6),
  Beautify (4), Annotate (3), Compute (1).
- **Domain:** Education & Examination (25), Business Operations (20), Other (18),
  Academic Papers (14), Engineering & Technology (10), Administrative Affairs (9),
  Financial Data (4).
- **Human labor time (min):** min 5, max 501, mean ≈ 139.
- **Task price proxy (yuan):** min 6, max 200, mean ≈ 47.
- **Price source:** `explicit_price` for 19 tasks, `estimated_price` for 81 tasks.

## Evaluation

OmegaUse-OfficeVal is evaluated with deterministic, code-based verifiers on the
final delivered artifacts rather than on a fixed execution trajectory. Agents may
use GUI actions, scripts, APIs, or hybrid strategies to complete a task. A task
score is computed in two stages: an artifact must first pass **all** usability
(`dim1`) items; its normalized completion score is then computed from the weighted
positive and negative task-completion (`dim2`) items, lower-bounded at zero.

## Data Construction & Privacy

All tasks originate from authentic workplace requests and were adapted through a
multi-stage pipeline: privacy-preserving instruction rewriting, expert review,
input-file reconstruction and manual de-identification, rubric generation with
expert revision, code-verifier generation, and human–code discrepancy resolution.
Sensitive and identifying information has been removed while preserving user
intent, constraints, and task-critical details.

## Citation

```bibtex
@inproceedings{omegause_officeval_2027,
  title     = {OmegaUse-OfficeVal: Benchmarking LLM Agents on Long-Horizon Office-Suite Tasks with Economic Grounding},
  booktitle = {Proceedings of the 2027 ACM SIGKDD International Conference on Knowledge Discovery and Data Mining (KDD '27)},
  year      = {2027}
}
```
