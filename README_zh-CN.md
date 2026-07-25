# Omegause Officeval

[English](README.md)

Omegause Officeval 是一个用于安全校验、批量执行和汇总 100 个 Office
文档评估器的 Python 框架。系统接收 ZIP 提交包，先完成结构和安全检查，
再在隔离子进程中运行 verifier，并输出 JSON 与 CSV 报告。

> 本仓库只包含评估框架和 verifier 源码，不分发基准文档或用户提交的
> Office 文件。

## 功能

- 检测 ZIP 路径穿越、加密、文件数量、解压体积和异常压缩比。
- 为 `officeval_001` 至 `officeval_100` 提供统一的
  `evaluate(directory: str) -> dict` 接口。
- verifier 独立子进程运行，并支持配置并发数和超时。
- 显示进度、当前运行编号、执行通道和耗时。
- 原子写入编号级 JSON、汇总 JSON 和 CSV。
- 96 个 verifier 可在 Windows、macOS 和 Linux 上使用普通模式。
- `011`、`023`、`039`、`081` 在 Windows 上使用串行 Office COM 通道。

## 环境要求

- Python 3.10 或更高版本。
- 普通模式支持 Windows、macOS 和 Linux。
- 只有需要运行四个 COM 硬依赖 verifier 时，才要求 Windows 安装
  Microsoft Word 和 Excel。

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

输入必须是 ZIP，根目录包含 100 个三位编号目录：

```text
officeval_001/
officeval_002/
...
officeval_100/
```

每个目录包含相应 verifier 所需的 Office 或 PDF 文件。预检只报告异常
内容，不会删除提交文件。

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

COM 模式只控制 `011`、`023`、`039`、`081`：

- `auto`：Windows 启用，macOS/Linux 跳过。
- `enabled`：要求 Windows 并启用这四项。
- `disabled`：所有平台跳过这四项。

存在静态解析路径的 verifier 始终使用普通模式，即使指定
`--com-mode enabled` 也不会启动 Office。

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
[Verifier 接口](docs/verifier-contract.md) 和
[安全模型](docs/security.md)。

## 参与贡献

提交 issue 或 pull request 前请阅读 [CONTRIBUTING.md](CONTRIBUTING.md)。
安全问题请按 [SECURITY.md](SECURITY.md) 私下报告，不要提交公开 issue。

## 许可证

本项目采用 [Apache License 2.0](LICENSE)，版权说明见 [NOTICE](NOTICE)。
依赖许可证见 [第三方许可证](THIRD_PARTY_LICENSES.md)。
