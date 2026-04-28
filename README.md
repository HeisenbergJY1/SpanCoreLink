# SpanCoreLink

**SAP2000 与 Rhino 的桥梁插件 | Bridge Plugin between SAP2000 and Rhino**

一键将 SAP2000 结构模型导入 Rhino，支持按组、截面、材料分图层显示，实体/线框自由切换。

Import SAP2000 structural models into Rhino with one click. Organize by group, section, or material with automatic layer coloring. Toggle between solid and wireframe views.

---

## 安装 | Installation

在 Rhino 的包管理器中搜索 **SpanCoreLink** 安装，或通过命令行：

Search for **SpanCoreLink** in Rhino's Package Manager, or via command line:

```
_PackageManager
```

### 依赖 | Dependencies

本插件依赖 [PySap2000](https://pypi.org/project/pysap2000/) 库与 SAP2000 通信。安装插件后，在 Rhino 中运行 `CheckUpdate` 命令可自动安装。

This plugin requires [PySap2000](https://pypi.org/project/pysap2000/) to communicate with SAP2000. After installing the plugin, run the `CheckUpdate` command in Rhino to install it automatically.

---

## 使用流程 | Workflow

```
1. 在 SAP2000 中打开并保存模型
2. 在 Rhino 中运行 GetSapModel   → 提取模型数据为 JSON
3. 运行任意导入命令               → 将模型导入 Rhino
```

```
1. Open and save your model in SAP2000
2. Run GetSapModel in Rhino        → Export model data to JSON
3. Run any import command          → Import the model into Rhino
```

---

## 命令列表 | Commands

| 命令 | 说明 |
|------|------|
| `GetSapModel` | 连接 SAP2000，自动同步单位，导出模型数据 |
| `ImportSapModel` | 按单元类型分图层导入（框架/索单元） |
| `ImportByGroup` | 弹窗选择组，按组分图层导入 |
| `ImportBySection` | 按截面分图层导入，支持标准化截面名 |
| `ImportByMaterial` | 按材料分图层导入 |
| `ForceEnglishInput` | 切换到 Rhino 时自动锁定英文输入法（Toggle） |
| `CheckUpdate` | 检查并升级 PySap2000 库 |
| `RestartRhino` | 保存文件并重启 Rhino |
| `ContactAuthor` | 查看教程链接与联系方式 |

---

## 功能特性 | Features

- **单位自动同步** — `GetSapModel` 会自动将 SAP2000 单位对齐到 Rhino 当前单位系统，导出后恢复原始单位
- **多种导入模式** — 实体（Solid）或线框（Wireframe），按需选择
- **自动分色** — 按组/截面/材料导入时，每个图层自动分配随机颜色，便于区分
- **截面标准化** — `ImportBySection` 支持将不同命名但规格相同的截面合并到同一图层
- **英文输入锁定** — 解决 Rhino 命令行误触中文输入法的痛点，切换窗口时自动恢复

---

## 环境要求 | Requirements

- Rhino 8
- SAP2000（需同时运行）
- Windows

---

## 开发 | Development

每个命令对应一个独立的 Python 文件，放在项目根目录。图标为 SVG 格式，放在 `icon/` 目录下。

Each command corresponds to a standalone Python file in the project root. Icons are SVG files in the `icon/` directory.

```
SpanCoreLink/
├── GetSapModel.py
├── ImportSapModel.py
├── ImportByGroup.py
├── ImportBySection.py
├── ImportByMaterial.py
├── ForceEnglishInput.py
├── CheckUpdate.py
├── RestartRhino.py
├── ContactAuthor.py
├── icon/
└── release/
```

---

## 作者 | Author

**jiangyao** · [spancore.cn](https://www.spancore.cn)

教程视频 | Tutorial: [【效率起飞】SAP2000模型一键导入Rhino！](https://www.bilibili.com/video/BV1v5Z9BREUe)

---

## 许可 | License

MIT
