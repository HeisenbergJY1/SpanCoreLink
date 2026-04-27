# SpanCoreLink 项目规则

## 禁止修改的文件
- `SpanCoreLink.rhproj` 文件禁止任何修改。该文件由 Rhino 项目编辑器管理，手动编辑会导致项目损坏。新增命令需要在 Rhino 编辑器中手动添加。

## 项目结构
- 每个命令对应一个独立的 Python 文件，放在项目根目录（如 `ImportSapModel.py`、`CheckUpdate.py`）
- 命令图标为 SVG 格式，放在 `icon/` 目录下，文件名与命令名一致（如 `icon/ImportSapModel.svg`）
- `__init__.py` 中的 `__all__` 列表需要包含所有命令模块名
- `release/` 目录存放发布产物，不要修改其中内容

## 代码规范
- 所有 Python 文件使用 UTF-8 编码，文件头加 `# -*- coding: utf-8 -*-`
- Rhino 环境检测使用 try/except 导入 `rhinoscriptsyntax`，设置 `RHINO_ENV` 标志
- UI 对话框使用 Eto.Forms（Rhino 内置的跨平台 UI 框架）
- 用户提示信息使用中文
- 每个命令文件底部包含 `if __name__ == "__main__":` 入口
