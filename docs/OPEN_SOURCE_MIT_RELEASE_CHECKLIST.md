# Open Source MIT Release Checklist

本清单用于把当前仓库整理为适合公开发布的 **MIT 开源版**。

> 目标：只公开你拥有权利的源码、配置与文档；不公开第三方规范原文、真实论文样本、测试产物和可识别个人 / 机构信息。

---

## 预检脚本（建议先执行）

在正式整理公开仓库前，先运行：

```powershell
.\check_public_release.bat

# 严格模式：发现阻塞项时返回非 0
.\check_public_release.bat --strict
```

说明：
- 该脚本会扫描高风险目录、生成产物、敏感提取文本、品牌图标占位等典型问题。
- 它不会自动删除文件，只负责定位，方便你在发布前人工确认与清理。

---

## 一、建议保留的内容

### 建议保留
- `src/`
- `app/`
- `main.py`
- `requirements.txt`
- `README.md`
- `LICENSE`
- `THIRD_PARTY_NOTICES.md`
- `CONTRIBUTING.md`
- 纯代码生成的测试
- 你本人原创、可再分发的最小 synthetic / public fixtures

### 可保留但需确认来源
- `src/ui/icons/app_icon.ico`
- `src/ui/icons/app_icon.png`
- `src/ui/icons/info.svg`
- `src/ui/icons/moon.svg`
- `src/ui/icons/sun.svg`

若图标不是你原创或无明确再分发许可，请先替换。

---

## 二、不要上传的内容

### A. 第三方规范原文、模板原件、提取文本
建议删除或移出公开仓库：
- `tests/former-templates/**`
- `tests/_scene_prompt_regression_*/**`
- `tmp_prompt_regression_exec_*/**`

尤其不要公开：
- `prompt_filled.txt`
- `spec_extracted.txt`
- `summary.json`
- `summary.md`
- `round_*.json`
- `round_*.log`
- `round_*_last_message.txt`

原因：这些文件可能直接包含第三方规范正文、摘录、中间提取结果或 prompt 产物，版权风险最高。

### B. 真实或疑似真实文档样本
建议删除或替换为 synthetic 样本：
- `tests/00/**`
- `tests/clone/**`
- `tests/test-1/**`
- `tests/test-equation/**`
- `tests/test-markdown/**`
- `tests/目录/**`

原因：
- 版权来源不明；
- 可能包含真实姓名、学校、单位、课题信息；
- 公开后难以证明你拥有再分发权利。

### C. 生成产物、对比稿、报告、临时目录
一律不要上传：
- 所有 `*_new.docx`
- 所有 `*_对比稿.docx`
- 所有 `*_报告.json`
- 所有 `*_报告.md`
- 所有 `*_排版附件/`
- 所有 `tmp_*.docx`
- 所有 `_tmp_*.docx`
- `e2e_*/`
- `exec-clone-check-*/`
- `scene-upgrade-check-*/`
- `scan_reports/`
- `tmp_md_audit/`
- `tmp_md_symbol_audit/`
- `tmp_toc_probe*/`

### D. 品牌图标与平台标识
建议确认来源后再公开：
- `src/ui/icons/github.svg`
- `src/ui/icons/bilibili.svg`

如果它们只是跳转按钮，更稳妥的方案是替换为自绘通用外链图标。

### E. 本地环境与构建目录
不要上传：
- `.venv/`
- `build/`
- `dist/`
- `.pytest_cache/`
- `__pycache__/`

---

## 三、公开前必须补齐的文件

### 1. LICENSE
保留 MIT 许可证文本。

### 2. THIRD_PARTY_NOTICES.md
说明：
- 本项目源码采用 MIT；
- 第三方依赖仍保留各自许可证；
- `PySide6` 不是 MIT，打包二进制时要额外核对其许可证义务。

### 3. README
建议明确写出：
- 公开仓库不包含第三方规范原文与真实论文样本；
- 测试样本应以 synthetic / public fixtures 为主；
- 发布前可运行公开发布扫描脚本。

### 4. CONTRIBUTING.md
建议明确写出：
- 不接收未授权的高校规范原文、论文原件、商业模板；
- 公开测试素材必须可合法再分发；
- 不确定来源时默认不要上传。

---

## 四、发布前自查

- [ ] 仓库中不含第三方规范原文或提取文本
- [ ] 仓库中不含真实论文或含个人 / 机构信息的样本文档
- [ ] 仓库中不含 `*_new.docx`、对比稿、报告和临时输出
- [ ] 图标、模板、封面素材来源明确
- [ ] `README.md`、`THIRD_PARTY_NOTICES.md`、`CONTRIBUTING.md` 已同步更新
- [ ] `.gitignore` 已覆盖常见本地产物与高风险样本
- [ ] 若分发 PyInstaller 二进制，已单独检查 Qt / PySide6 许可要求

---

## 五、建议的公开仓库结构

```text
app/
src/
tests/
  fixtures_public/
  test_*.py
docs/
README.md
LICENSE
THIRD_PARTY_NOTICES.md
CONTRIBUTING.md
requirements.txt
```

---

## 六、默认原则

如果你不能明确证明某个文件：
- 是你本人原创；
- 不含第三方规范正文；
- 不含真实个人 / 机构信息；
- 你有权公开再分发；

那么就**默认不要上传**。
