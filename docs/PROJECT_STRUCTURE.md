# PROJECT_STRUCTURE

## 当前产品定位
- `0.20 LTS`：毕业论文格式专项修改工具。
- 当前发布不以“多场景通用排版平台”为目标。
- 其他场景拟在 `1.00` 版本完成开发与发布。

## 目录概览
- `app/`：应用入口与启动包装。
- `src/`：核心实现。
  - `src/ui/`：桌面界面。
  - `src/engine/`：排版流水线与规则。
  - `src/scene/`：场景模板加载、保存、迁移。
  - `src/docx_io/`：DOCX 读写、对比、域刷新。
  - `src/formula_core/`：公式解析、修复、转换。
  - `src/report/`：JSON / Markdown 报告输出。
  - `src/utils/`：公共工具与元数据。
- `scripts/windows/`：Windows 安装、启动、打包脚本。
- `templates/`：源码模式下的用户可写模板目录。
- `tests/`：自动化测试与部分回归样例。
- `docs/`：项目文档。

## 当前发布入口
- 启动入口：`main.py`
- GUI 入口：`app/main.py`
- 发布脚本：`scripts/windows/package_release.bat`
- 当前发布 spec：`Lark-Formatter_v0.20_LTS.spec`

## 当前发布产物
- 目录：`dist/Lark-Formatter_v0.20_LTS/`
- 可执行文件：`dist/Lark-Formatter_v0.20_LTS/Lark-Formatter.exe`
