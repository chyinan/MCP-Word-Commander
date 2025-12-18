# MCP_WORD_EDITER

轻量级的文本处理与编辑工具（开发中）。

## 简介
MCP_WORD_EDITER 用于在本地对文本/词条进行快速处理与编辑，基于 Python 实现的简单服务。

## 依赖
- Python 3.8+
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