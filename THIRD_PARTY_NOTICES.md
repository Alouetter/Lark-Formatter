# Third-Party Notices

本项目源码采用 **MIT License** 发布。  
第三方依赖及其许可证**不因本项目采用 MIT 而改变**；使用、分发或打包时仍需遵守各依赖自身许可证。

> 下列信息根据当前本地开发环境中的包元数据整理。

## Runtime / build dependencies

| Dependency | Version seen locally | License |
|---|---:|---|
| python-docx | 1.2.0 | MIT |
| PySide6 | 6.10.2 | LGPL-3.0-only OR GPL-2.0-only OR GPL-3.0-only |
| lxml | 6.0.2 | BSD-3-Clause |
| latex2mathml | 3.79.0 | MIT |
| olefile | 0.47 | BSD |
| pywin32 | 311 | PSF |

## Important note on PySide6 / Qt

`PySide6` **不是 MIT**。  
如果你只是公开源码，通常只需在仓库中保留第三方许可证说明即可。  
如果你进一步分发：

- PyInstaller 打包后的桌面应用；
- 安装包；
- 带 Qt 运行时的二进制发布物；

则需要**单独检查并满足 Qt / PySide6 对应许可证要求**。

## Project policy

- 本项目自身代码：MIT
- 第三方依赖：各自原许可证
- 本仓库不试图重新许可任何第三方库、第三方模板、第三方规范原文或第三方文档样本

## Recommended practice

- 发布源码仓库时，保留本文件
- 发布打包二进制时，同时附带相应第三方许可证文本
- 如后续新增依赖，请同步更新本文件
