# Excel 小表并大表 — 智能合并工具V1.0

[![AI Vibe-Coding](https://img.shields.io/badge/AI_Vibe--Coding-🤖-purple)](https://github.com/OrangeMoon-hub/Excel_combiner)
[![Python](https://img.shields.io/badge/Python-3.8+-blue.svg)](https://www.python.org/)
[![Zero Dependencies](https://img.shields.io/badge/Dependencies-Zero-brightgreen.svg)](#零第三方依赖)
[![Tested](https://img.shields.io/badge/Tested-19_Cases_+_5_E2E_Rounds-success.svg)](#验证记录)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

> 🤖 **这是一个 AI Vibe-Coding 项目。** 全部代码由 AI 助手生成，橙月君（数据岗位）负责需求定义、验收测试与业务场景验证。
>
> 双击运行、图形化对话框、零依赖。将多个结构相似的 Excel/CSV 数据文件按列名智能匹配合并到统一的模板表中。
> 目前项目仍在增加功能，故没有编译。各位访客可以下载“测试环境”文件夹中的excel和py脚本，放置在同一目录下运行测试；

---

## ✨ 为什么用这个？

手动合并几十个仓库/部门交上来的 Excel，逐个打开复制粘贴——**费时、易错、不可审计**。

这个工具让你：

| 痛点 | 解决 |
|------|------|
| 列名不一致 | 自动匹配 + 检测异常（丢失列/多余列/同名列冲突） |
| 格式难统一 | 以模板表为基准，严格按列对齐 |
| 出错了不知道 | 生成完整异常记录表，逐条可追溯 |
| 不会写代码 | 双击 exe，图形化对话框，零编程门槛 |
| 装环境麻烦 | 零第三方依赖，仅用 Python 标准库，单文件 exe |

---

## 🖥️ 运行方式

```
双击 合并脚本.exe
      │
      ▼
  对话框一：选择模板文件（单选）
      │
      ▼
  对话框二：选择待合并文件（多选，支持全选/反选）
      │
      ▼
  逐文件处理：遇到异常弹窗询问 → 用户决策
      │
      ▼
  生成 合并结果.xlsx
```

---

## 📋 功能一览

### 智能文件管理
- 自动扫描目录内所有 `.xlsx` / `.csv` 文件
- 图形化单选模板 + 多选待合并文件（勾选变色 + 全选按钮）
- 排除 `~$` 临时锁文件

### 列名匹配 & 冲突处理

| 场景 | 行为 |
|------|------|
| 列名完全一致 | 自动匹配，按列填值 |
| **同名列**（小表多列同名） | 三步判断：值一致→全填 / 数量相等→顺序 / 多于→逐列弹窗 |
| 列名不存在于模板 | 弹窗确认后丢弃，记录异常 |
| 数据列数超过表头 | 弹窗确认后截断，逐额外列记录 |
| 表头有空列名 | 弹窗确认后丢弃，标注"数据异常，无列名" |

### 结果输出

| Sheet | 内容 |
|-------|------|
| 模板 Sheet(s) | 模板表头 + 所有合并数据行 |
| `异常记录` | 四种异常类型的完整记录（含前 5 行数据值、列位置、用户操作） |
| `取消合并的表名` | 被取消合并的文件及原因 |

### 异常记录样例

```
文件名     Sheet    行号    列名             列1位置  列1值        异常类型
小表3.xlsx  NAME1   全部行   数据异常，无列名    C       v1,v2,v3    无列名
小表3.xlsx  NAME1   全部行   -                 I       45,67,AA;    超表头列数
小表2.xlsx  NAME2   -       111               L       111,222       列名不存在于大表
小表1.xlsx  NAME1   -       T                 K       落选值1,2     同名列未选择
```

---

## ✅ 验证记录

> 本项目由 **数据岗位人员（橙月君）** 亲自验收每一步产出，确保结果可信赖。

| 阶段 | 验证内容 | 结果 |
|------|---------|------|
| **单元测试** | 同名列三步判断（值一致→全填 / 数量相等→顺序 / 多于→弹窗）4 项 | ✅ 全部通过 |
| **测试方案** | 19 项结构化测试用例，覆盖基本合并、同名列冲突、异常列处理、Sheet 匹配、特殊字符、多文件批量、取消合并等 | ✅ 全部通过 |
| **E2E 第 1 轮** | 原始版本在 Windows 实机运行，修复 Excel 修复弹窗问题 | ✅ |
| **E2E 第 2 轮** | 第二轮深挖 XML 命名空间覆盖问题（ns0 前缀） | ✅ |
| **E2E 第 3 轮** | 同名列升级三步判断 + 异常记录信息密度提升（补前 5 行值） | ✅ |
| **E2E 第 4 轮** | 图形化文件选择对话框 + CSV 支持 + 跨平台 tkinter 兼容 | ✅ |
| **E2E 第 5 轮** | 最终端到端：Windows + Linux 双平台，实际业务数据跑通 | ✅ |
| **手动验收** | 橙月君以真实数据岗位场景实测，逐功能点核对输出结果 | ✅ 可信赖 |

> 共计 **19 个测试用例** 覆盖 6 大场景维度，**5 轮端到端回归**，**双平台（Windows/Linux）** 验证通过。

---

## 🛠️ 技术栈

| 维度 | 选择 |
|------|------|
| 语言 | Python 3.8+ |
| GUI | tkinter（标准库） |
| Excel 读写 | zipfile + xml.etree（标准库，纯 Python 解析/写入 xlsx） |
| CSV 读取 | csv（标准库） |
| 打包 | PyInstaller `--onefile` |
| 依赖 | **零第三方库** |

---

## 📦 项目结构

```
Excel_combiner/
├── 合并脚本.py                        # 主程序（零依赖）
├── PRD_Excel小表并大表工具.md          # 产品需求文档
├── 测试文档/                          # 测试用例 & 文档
├── 测试环境/                          # 测试用 Excel 文件
├── .gitignore
├── LICENSE
└── README.md
```

---

## 🚀 本地运行 / 打包

### 直接运行

```bash
python 合并脚本.py
```

### 打包为 exe

```bash
pip install pyinstaller
pyinstaller --onefile --console 合并脚本.py
```

> 使用 `--console` 保留控制台窗口，方便查看进度和日志。

---

## ⚠️ 约束 & 边界

- 模板文件第 1 列必须为 `表名`（合并时自动填入来源文件名）
- Excel 仅支持 `.xlsx` 格式（不支持 `.xls`）
- CSV 以 UTF-8 读取，映射到虚拟 Sheet `Sheet1`，仅当模板含同名 Sheet 时匹配
- Windows tkinter 不支持 `pady=tuple` 写法（`pady=(15,5)` 报 `bad screen distance`），间距必须用单整数

完整边界情况见 [PRD 第 5 节](PRD_Excel小表并大表工具.md#5-边界情况与约束)。

---

## 👤 作者

**OrangeMoon** (橙月君)

- GitHub: [@OrangeMoon-hub](https://github.com/OrangeMoon-hub)

---

## 📄 License

MIT © OrangeMoon
