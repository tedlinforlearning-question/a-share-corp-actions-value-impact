# A-Share Corporate Actions — Valuation Impact Toolkit

> DCF valuation for IPOs, M&A, and refinancing. A practical investment-banking skill for WorkBuddy / CodeBuddy.

---

## What it does

Note: The project is developed primarily in Chinese, as it is designed specifically for the A-share market, including PRC regulatory requirements and local valuation practices.

Covers six scenarios, with a guided Q&A decision tree:

| Scenario | What is valued | Core method |
|:---|:---|:---|
| **IPO** | Capital-raising projects (incremental target) | Standalone DCF + project FCFF → pro-forma |
| **M&A — acquisition** | The target company | DCF + target FCFF + synergies → pro-forma |
| **M&A — disposal** | The divested subsidiary | Disposal-type test → residual DCF |
| **Refinancing — WC only** | Dilution impact | Ownership / BPS / EPS + interest savings |
| **Refinancing — WC + projects** | Dilution + project DCF | Dilution + project DCF → combined |
| **Refinancing — acquisition** | Dilution + target | Dilution + target DCF + pro-forma |

Key capabilities: PDF parsing · construction-in-progress cross-check · non-recurring separation · dilution analysis · comparable-company WACC (Wind / AKShare / manual) · full DCF · auto-generated Excel.

---

## Quick start

```bash
pip install PyMuPDF openpyxl
```

| Data source | How | For |
|:---|:---|:---|
| 🔵 **Wind** | API key at https://aifinmarket.wind.com.cn/#/docs | IB teams with a Wind terminal |
| 🟢 **AKShare** (free) | `pip install akshare` | Individuals |
| ⭐ **Manual input** | Provide β / WACC / growth directly | Zero dependency |

> Core analysis works without any cloud data; only WACC calibration needs external inputs.

---

## Install

**Option 1 — WorkBuddy (recommended):** Download this repo (Code → Download ZIP), unzip, then Skills → Add skill → Upload. Say **"help me value an IPO"** or **"help me value an M&A deal"**.

**Option 2 — Excel templates:** Download the `.xlsx` files under `excel_templates/` and model manually.

**Option 3 — Python scripts:**

```bash
python scripts/build_dcf_model.py --all
python scripts/extract_financials.py prospectus.pdf
```

---

## Project structure

```
├── SKILL.md                # Core skill definition
├── README.md               # This file
├── docs/                   # Install / usage / limitations
├── excel_templates/        # Excel templates
├── scripts/                # Helper scripts
└── references/             # Reference parameters (anonymized)
```

---

## License

MIT License

---

# A股资本运作 · 价值影响建模 Skill（中文）

> 面向 IPO、并购重组与再融资的 DCF 估值建模，投行实务级，可用于 WorkBuddy / CodeBuddy。

---

## 功能概述

覆盖六大场景，交互式问答决策树引导：

| 场景 | 评估标的 | 核心方法 |
|:---|:---|:---|
| **IPO** | 募投项目（增量标的） | 独立DCF + 募投FCFF → 备考合并 |
| **M&A 并购置入** | 标的公司 | DCF + 标的FCFF + 协同 → 备考合并 |
| **M&A 重大资产出售** | 剥离并表子公司 | 出售类型判定 → 剩余业务DCF |
| **再融资-纯补流** | 摊薄效应 | 持股/BPS/EPS + 利息节约 |
| **再融资-补流+募投** | 摊薄 + 募投项目 | 摊薄 + 项目DCF → 合并 |
| **再融资-收购+配套** | 摊薄 + 标的公司 | 摊薄 + 标的DCF + 备考合并 |

核心能力：PDF 智能解析 · 在建工程/募投交叉比对 · 备考扣非分离 · 再融资摊薄分析 · 可比公司 WACC（Wind / AKShare / 手动输入）· 完整 DCF 建模 · Excel 自动生成。

---

## 快速开始

```bash
pip install PyMuPDF openpyxl
```

| 数据源 | 接入方式 | 适合 |
|:---|:---|:---|
| 🔵 **Wind 万得** | https://aifinmarket.wind.com.cn/#/docs 获取 API Key | 有 Wind 终端的团队 |
| 🟢 **AKShare**（免费） | `pip install akshare` | 个人用户 |
| ⭐ **手动输入** | 直接提供 β/WACC/增速 | 零依赖 |

> 核心分析无需云端数据，仅 WACC 校准和可比公司环节需要外部数据，可手动填入。

---

## 安装使用

**方式一 WorkBuddy（推荐）：** 下载本仓库（Code → Download ZIP）解压 → 技能管理 → 添加技能 → 上传文件夹 → 对话中说 **"帮我做 IPO 估值"** 或 **"帮我做并购估值"**。

**方式二 Excel 模板：** 直接下载 `excel_templates/` 下的模板手动建模。

**方式三 Python 脚本：**

```bash
python scripts/build_dcf_model.py --all
python scripts/extract_financials.py 招股书.pdf
```

---

## 目录结构

```
├── SKILL.md                # 核心 Skill 配置
├── README.md               # 本文件
├── docs/                   # 安装/使用/局限性说明
├── excel_templates/        # Excel 模板
├── scripts/                # 辅助脚本
└── references/             # 案例参数（脱敏版）
```

## 许可

MIT License
