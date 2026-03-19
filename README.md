# Lark-Formatter

毕业论文 `docx` 格式专项修订桌面工具（PySide6 + python-docx）。

## 当前版本
- `0.20 LTS`
- 当前版本**仅面向毕业论文格式调整场景**。
- 其他场景（如更广义的通用文档 / 扩展模板体系）计划在 `1.00` 版本完成开发与发布。

## 版本定位
- `0.20 LTS`：毕业论文格式专项工具，聚焦论文排版修订、格式统一与审计输出。
- `1.00`（规划中）：扩展到其余场景与更完整的多模板能力。

## 核心能力
- 毕业论文专项排版：支持论文模板加载、另存、重命名、删除。
- 规则流水线：页面设置、样式统一、标题识别与编号、目录处理、图表题注、表格处理、分节格式化、页眉页脚、校验。
- 实验室功能：Markdown 清理、空白与全半角规范（基础空白清洗 + 语境字符转换）、正文引用域关联、化学式上下角标恢复、公式表格识别调整。
- 多产物输出：排版后文档、对比稿、JSON / Markdown 报告。
- 模板克隆：可从参考 `docx` 克隆样式 / 页面 / 标题编号为新的**论文模板**。

## 适用范围
- 适用：毕业论文 / 学位论文格式修订与校对辅助。
- 当前不建议作为：公文排版工具、通用文档排版工具、跨行业多场景模板平台。

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
- `dist/Lark-Formatter_v0.20_LTS/`

## 打包后的模板目录说明（重点）
在封包运行时会同时看到两个 `templates`：
- `dist/Lark-Formatter_v0.20_LTS/templates`：**用户可读写目录（实际读写路径）**
- `dist/Lark-Formatter_v0.20_LTS/_internal/templates`：内置资源目录（只读种子）

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

## 开源发布与合规
- 项目自身源码按 `MIT` 发布，第三方依赖仍保持各自许可证，见 `THIRD_PARTY_NOTICES.md`。
- 公开仓库建议只保留源码、原创文档与可再分发的 synthetic / public fixtures。
- 公开前可运行检查脚本：

```powershell
# 普通扫描
.\check_public_release.bat

# 严格模式：发现阻塞项时返回非 0，适合发布前自检 / CI
.\check_public_release.bat --strict
```

## 相关文档
- `docs/PROJECT_STRUCTURE.md`
- `docs/RELEASE_NOTES_0.20_LTS.md`
- `docs/OPEN_SOURCE_MIT_RELEASE_CHECKLIST.md`
- `docs/NOT_FOR_UPLOAD.md`
- `THIRD_PARTY_NOTICES.md`
- `CONTRIBUTING.md`
- `tests/fixtures_public/README.md`
- `使用说明.md`
