# OmegaUse-OfficeVal

[English](README.md)

<p align="center">
  <a href="https://omegause-officeval.github.io/"><strong>项目网站</strong></a> &nbsp;•&nbsp;
  <a href="https://huggingface.co/datasets/baidu-frontier-research/OmegaUse-OfficeVal"><strong>Hugging Face 数据集</strong></a> &nbsp;•&nbsp;
  <a href="https://arxiv.org/abs/XXXX.XXXXX"><strong>论文（即将发布）</strong></a>
  <br>
  <a href="https://github.com/baidu-frontier-research/OmegaUse-OfficeVal"><strong>源码</strong></a> &nbsp;•&nbsp;
  <a href="https://github.com/baidu-frontier-research/OmegaUse-OfficeVal/issues"><strong>问题反馈</strong></a> &nbsp;•&nbsp;
  <a href="https://github.com/baidu-frontier-research/OmegaUse-OfficeVal/releases"><strong>版本发布</strong></a>
</p>


OmegaUse-OfficeVal 是一个用于安全校验、批量执行和汇总 100 个 Office
文档评估器的 Python 框架。系统接收 ZIP 提交包，先完成结构和安全检查，
再在隔离子进程中运行 verifier，并输出 JSON 与 CSV 报告。

> 本仓库包含评估框架和 verifier 源码；基准数据通过上方 Dataset 单独发布，
> 不分发用户提交文件或评估产生的工作目录。


## 基准框架

OmegaUse-OfficeVal 将真实长周期 Office 任务采集、经济价值估算与迭代式
代码验证结合起来，并为脱敏后的任务指令和输入文件配套细粒度评分规则及
可执行 verifier。

<p align="center">
  <img src="assets/benchmark-framework.png" alt="OmegaUse-OfficeVal 基准框架" width="100%">
</p>

## 功能

- 检测 ZIP 路径穿越、加密、文件数量、解压体积和异常压缩比。
- 为 `officeval_001` 至 `officeval_100` 提供统一的
  `evaluate(directory: str) -> dict` 接口。
- verifier 独立子进程运行，并支持配置并发数和超时。
- 显示进度、当前运行编号、执行通道和耗时。
- 原子写入编号级 JSON、汇总 JSON 和 CSV。
- 91 个 verifier 可在 Windows、macOS 和 Linux 上使用普通模式。
- `001`、`008`、`019`、`022`、`023`、`030`、`039`、`074`、`081`
  在 Windows 上使用串行 Office COM 通道。


## 环境要求

- Python 3.10 或更高版本。
- 普通模式支持 Windows、macOS 和 Linux。
- 只有需要运行九个 COM 硬依赖 verifier 时，才要求 Windows 安装
  Microsoft Office。

## 平台兼容性

| 平台 | 普通模式 | Office COM | 持续集成状态 |
| --- | --- | --- | --- |
| Windows | 支持 | 安装 Microsoft Office 后支持九个指定 verifier | 已在 Python 3.10 和 3.12 上验证 |
| Linux | 支持 | 不支持；`auto` 模式会跳过 COM 强依赖项 | 已在 Ubuntu、Python 3.10 和 3.12 上验证 |
| macOS | 预期支持 | 不支持；`auto` 模式会跳过 COM 强依赖项 | 当前未纳入 CI |

普通模式依赖静态文档解析；平台相关的 Office 渲染和 COM 自动化仅在
Windows 上可用。

## 安装


```bash
python -m venv .venv
```

Linux 或 macOS：

```bash
source .venv/bin/activate
python -m pip install -e ".[test]"
```

Windows PowerShell：

```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install -e ".[test]"
```

`pywin32` 只会在 Windows 上安装。

## 提交包结构

完整提交包应为 ZIP，并在根目录包含 `officeval_001/` 至
`officeval_100/` 共 100 个任务目录：

```text
officeval_001/
officeval_002/
...
officeval_100/
```

缺少编号目录会产生 Warning，而不是压缩包级 Fatal。用户确认继续后，
缺失目录、空目录或目录中没有受支持文档的编号不会进入 verifier，系统会
生成正常的维度一失败结果，总分和完成度均为 0。

支持的文档扩展名：

- Word：`.docx`
- Excel：`.xlsx`、`.xlsm`
- PowerPoint：`.pptx`
- PDF：`.pdf`

不支持旧版 Office 格式 `.doc`、`.xls`、`.ppt`。这些格式依赖平台相关
转换组件，无法在 Windows、macOS 和 Linux 上提供一致行为。预检只报告
异常内容，不会删除提交文件。


## 使用

```bash
omegause-officeval --package /absolute/path/to/submission.zip
```

也可以使用模块入口：

```bash
python -m omegause_officeval --package /absolute/path/to/submission.zip
python -m core --package /absolute/path/to/submission.zip
```

常用参数：

```text
--max-workers N
--timeout-seconds SECONDS
--com-mode auto|enabled|disabled
```

COM 模式控制 `001`、`008`、`019`、`022`、`023`、`030`、`039`、
`074`、`081`：

- `auto`：Windows 启用，macOS/Linux 跳过。
- `enabled`：要求 Windows 并启用这九项。
- `disabled`：所有平台跳过这九项。

`011` 保留受控 COM fallback，但调度器始终以普通模式运行并禁用该
fallback。其他 verifier 不会启动 Office COM。


预检完成后，CLI 会列出 Fatal 和 Warning，并要求明确确认后才开始评估。

## 输出

每次提交在 `results/` 下生成独立任务目录：

```text
job.json
validation_report.json
summary.json
summary.csv
details.csv
001.json
...
100.json
```

`results/`、`submissions/` 和 `workspaces/` 是本地运行状态，不纳入版本控制。
字段定义、状态语义、完成度计算和缺失交付处理见
[结果格式](docs/result-format.md)。


## 清理工作目录

评估完成后，解压内容会保留在 `workspaces/<job_id>/`，便于复核。可使用独立清理工具释放空间：

```bash
# 只列出可清理任务、状态、修改时间和预计释放空间，不删除文件
python -m core.cleanup --list

# 清理指定任务的工作目录
python -m core.cleanup --job-id "<job_id>"

# 清理超过 30 天且已结束的工作目录
python -m core.cleanup --older-than-days 30
```

删除命令会先显示候选目录并要求输入 `y` 或 `yes` 确认。工具只删除 `workspaces/<job_id>/`；`submissions/<job_id>/` 中的原始 ZIP 和 `results/<job_id>/` 中的 JSON、CSV 结果会保留。运行中任务、状态未知任务、符号链接或目录联接不会被清理。

## 开发验证

```bash
python -m compileall -q core verifiers omegause_officeval
python -m pytest
python -m build
```

设计说明见 [架构](docs/architecture.md)、
[Verifier 接口](docs/verifier-contract.md)、
[结果格式](docs/result-format.md) 和
[安全模型](docs/security.md)。


## 参与贡献

提交 issue 或 pull request 前请阅读 [CONTRIBUTING.md](CONTRIBUTING.md)。
安全问题请按 [SECURITY.md](SECURITY.md) 私下报告，不要提交公开 issue。

## 许可证

本项目采用 [Apache License 2.0](LICENSE)，版权说明见 [NOTICE](NOTICE)。
依赖许可证见 [第三方许可证](THIRD_PARTY_LICENSES.md)。
