# Lark-Formatter

论文 `docx` 一键排版桌面工具（PySide6 + python-docx）。

## 核心能力
- 场景化排版：支持预设模板加载、另存、重命名、删除。
- 规则流水线：页面设置、样式统一、标题识别与编号、目录处理、图表题注、表格处理、分节格式化、页眉页脚、校验。
- 实验室功能：Markdown 清理、空白字符统一、正文引用域关联、化学式上下角标恢复、公式表格识别调整。
- 多产物输出：排版后文档、对比稿、JSON/Markdown 报告。
- 模板克隆：可从参考 `docx` 克隆样式/页面/标题编号为新模板。

## 环境要求
- Windows（脚本为 `.bat`，封包目标为 Windows）
- Python 3.10+

## 快速开始
```powershell
# 1) 首次安装依赖
.\install_env.bat

# 2) 启动 GUI
.\start_app.bat

# 3) 调试模式启动（有控制台输出）
.\start_app_debug.bat
```

## 打包发布
```powershell
.\package_release.bat
```

打包输出目录：
- `dist/Lark-Formatter_v0.1.0/`

## 打包后的模板目录说明（重点）
在封包运行时会同时看到两个 `templates`：
- `dist/Lark-Formatter_v0.1.0/templates`：**用户可读写目录（实际读写路径）**
- `dist/Lark-Formatter_v0.1.0/_internal/templates`：内置资源目录（只读种子）

启动时会把内置模板补齐到可写目录（仅补缺，不覆盖用户文件）。

模板目录解析优先级：
1. 若设置环境变量 `DOCX_FORMATTER_TEMPLATES_DIR`，优先使用该目录。
2. 封包模式使用 `Lark-Formatter.exe` 同级的 `templates`。
3. 源码模式优先使用独立 `templates` 目录（若存在），否则回退到 `src/scene/presets`。

## 常用环境变量
- `DOCX_FORMATTER_TEMPLATES_DIR`：强制指定模板读写目录。
- `DOCX_PIPELINE_STRICT_MODE`：流水线严格模式（`1/true/on` 开启）。
- `DOCX_DISABLE_FIELD_REFRESH=1`：禁用 Word 域刷新。
- `DOCX_FIELD_REFRESH_TIMEOUT_SEC`：Word 域刷新超时秒数（默认 10）。

## 开发与测试
```powershell
# 运行测试
.\.venv\Scripts\python.exe -m pytest -q

# 语法检查示例
.\.venv\Scripts\python.exe -m py_compile src\scene\manager.py
```

## 项目结构
- `app/main.py`：应用入口。
- `src/`：核心代码（UI、规则引擎、场景配置、docx 读写、报告）。
- `scripts/windows/`：Windows 启动、安装、打包脚本。
- `docs/`：项目文档。
- `main.py` 与根目录批处理：兼容入口。

## 相关文档
- `docs/PROJECT_STRUCTURE.md`
- `使用说明.md`
