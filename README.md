# MCP_WORD_EDITER

轻量级的文本处理与编辑工具。

## 简介
MCP_WORD_EDITER 用于在本地对文本/词条进行快速处理与编辑，基于 Python 实现的简单服务。

## 依赖
- Python 3.10+
- 见 `requirements.txt`

## 快速开始
```bash
python -m venv venv
venv\Scripts\activate    # Windows
# 或 (macOS / Linux)
# source venv/bin/activate
pip install -r requirements.txt
python server.py
```

服务器默认在 http://localhost:8000 （如有需要请在 `server.py` 中查看与修改）。

## 项目结构（示例）
- `server.py` - 启动脚本
- `requirements.txt` - Python 依赖

## 贡献
欢迎开 issue 或 PR。提交前请确保代码风格一致并包含必要说明。

## 许可证
本项目使用 MIT 许可证，详见 `LICENSE`。

## English

# MCP_WORD_EDITOR

Lightweight text processing and editing tool.

## Overview
MCP_WORD_EDITOR is a simple local service implemented in Python for quickly processing and editing text/entries.

## Requirements
- Python 3.10+
- See `requirements.txt`

## Quick Start
```bash
python -m venv venv
venv\Scripts\activate    # Windows
# or (macOS / Linux)
# source venv/bin/activate
pip install -r requirements.txt
python server.py
```

The server runs by default at http://localhost:8000 (check or modify `server.py` if needed).

## Project Structure (example)
- `server.py` - startup script
- `requirements.txt` - Python dependencies

## Contributing
Issues and pull requests are welcome. Please ensure code style consistency and include necessary descriptions before submitting.

## License
This project is licensed under the MIT License. See `LICENSE` for details.