# A股IPO · 重组 · 再融资 DCF 估值建模 Skill

> 投行实务级 IPO、并购重组与再融资 DCF 估值建模工作流，可直接在 WorkBuddy / CodeBuddy 中使用。

[![Version](https://img.shields.io/badge/version-1.3.0-blue)](./SKILL.md)

---

## 概述

本 Skill 提供一套完整的投行估值建模工作流，覆盖六大核心场景：

| 场景 | 评估标的 | 核心方法 |
|:---|:---|:---|
| **IPO 估值** | 募投项目（作为增量标的） | 独立DCF + 募投FCFF → 备考合并DCF |
| **M&A 并购置入** | 标的公司 | 独立DCF + 标的FCFF + 协同效应 → 备考合并 |
| **M&A 重大资产出售** | 剥离并表子公司 | 出售类型判定 → 备考扣非分离 → 剩余业务DCF |
| **再融资-纯补流** | 摊薄效应 | 持股/BPS/EPS变化 + 利息节约 + Pro Forma |
| **再融资-补流+募投** | 摊薄 + 募投项目 | 摊薄分析 + 项目DCF → 合并估值 |
| **再融资-收购+配套** | 摊薄 + 标的公司 | 摊薄分析 + 标的DCF + 备考合并 + 协同量化 |

### 核心能力

- 📄 **报告智能解析**：自动读取 PDF 年报、招股书、审计报告、重组报告书、募集说明书
- 📊 **在建工程/募投分析**：交叉比对募投项目与在建工程，修正投产时间线
- 🏷️ **出售类型判定**：自动识别出售标的是否适合 DCF 建模
- 🔍 **备考扣非分离**：区分经常性/非经常性损益，确保 DCF 基于真实经营状况
- 📉 **再融资摊薄分析**：发行前后持股/BPS/EPS/利息节约/Pro Forma 全套测算，覆盖纯补流、补流+募投、收购+配套三大子场景
- 📈 **可比公司分析**：βU 去杠杆矩阵 + WACC 计算（支持 Wind/AKShare/手动输入）
- 🧮 **完整 DCF 建模**：独立DCF → 募投/标的FCFF → 备考合并 → 敏感性分析
- 📋 **Excel 自动生成**：专业投行格式，含样式规范、假设清单、敏感性矩阵
- ⚠️ **业绩承诺可行性分析**（v1.3 新增）：压力测试 + 增长驱动力拆解 + 触发概率判断
- 💰 **配套募资子表**（v1.3 新增）：双重摊薄 + 募投项目独立DCF + 含募投合并调整

---

## 快速开始（约 15 分钟）

> 📖 完整图文指南见 [`docs/install_guide.md`](docs/install_guide.md)

### 第一步：安装 Python 环境

```bash
# 1. 确保 Python 3.9+ 已安装
python --version

# 2. 安装必需包（PDF 读取 + Excel 生成）
pip install PyMuPDF openpyxl
```

### 第二步：接入金融数据源（可选）

| 数据源 | 如何接入 | 适合谁 |
|:---|:---|:---|
| 🔵 **Wind 万得** | 在 https://aifinmarket.wind.com.cn/#/docs 获取 API Key，安装 `wind-mcp-skill` | 公司有 Wind 终端的投行团队 |
| 🟢 **AKShare**（免费） | `pip install akshare`，无需注册 | 没有 Wind 的个人用户 |
| ⭐ **手动输入** | 直接提供 β值/WACC/行业增速等参数 | 零依赖，任意环境可用 |

> 💡 **无云端数据源也能使用**：核心分析（DCF建模、摊薄分析、备考合并）不需要云端数据。仅 WACC 校准和可比公司环节需要外部数据，可手动填入。

## 🚀 安装使用

### 方式一：WorkBuddy 上传（推荐）

1. 下载本仓库（点击 Code → Download ZIP，解压）
2. 在 WorkBuddy 中打开技能管理 → **添加技能** → **上传技能**
3. 选择解压后的文件夹，系统自动识别 `SKILL.md` 完成配置
4. 对话中说 **"帮我做 IPO 估值"** 或 **"帮我做并购估值"** 即可开始

### 方式二：直接使用模板

不需要 WorkBuddy？直接下载 [`excel_templates/`](excel_templates/) 目录下的 Excel 模板手动建模。

### 方式三：Python 脚本独立运行

```bash
pip install PyMuPDF openpyxl akshare
python scripts/build_dcf_model.py --all   # 生成 Excel
python scripts/extract_financials.py 招股书.pdf  # 提取 PDF 数据
```



### 使用流程

#### IPO 项目

```
提供文件 → 前置检查 → 在建工程分析 → 可比公司WACC
→ 独立DCF → 募投FCFF → 备考合并DCF → 敏感性分析
```

**所需报告：**
- 招股说明书（申报稿/注册稿）
- 审计报告（申报稿）

#### M&A 并购置入项目

```
提供文件 → 前置检查 → 可比公司 → 独立DCF
→ 业绩承诺检查（条件启用）→ 标的FCFF + 协同效应
→ 配套募资分析（条件启用）→ 备考合并DCF → 敏感性分析
```

**所需报告：**
- 上市公司近3年年报
- 重组报告书/重大资产购买报告书
- 标的公司审计报告/模拟报表

#### M&A 重大资产出售项目

```
提供文件 → 出售类型判定 → 备考扣非分离
→ 出售前DCF判定 → 剩余业务WACC → 剩余业务DCF → 备考对比 → 敏感性
```

**所需报告：**
- 上市公司近3年年报
- 重组报告书（含"重大资产出售"）
- 置出资产专项审计报告
- 备考审阅报告
- 资产评估报告

#### 再融资-纯补流项目

```
提供文件 → 前置检查 → 参数提取 → 摊薄分析(7 Sheet Excel) → 最终判断
```

**所需报告：**
- 募集说明书（注册稿/申报稿）
- 审计报告（或年报全文PDF，用于提取BPS/资产负债率）

#### 再融资-补流+募投项目

```
提供文件 → 前置检查 → 参数提取 → 摊薄分析(7 Sheet) + 项目DCF(3 Sheet) → 合并估值
```

**所需报告：**
- 募集说明书（含募投项目效益测算表）
- 审计报告（或年报全文PDF）

#### 再融资-收购+配套项目

```
提供文件 → 前置检查 → 参数提取 → 摊薄分析(7 Sheet) + 标的DCF + 备考合并 + 协同量化
```

**所需报告：**
- 募集说明书（含收购方案+发行方案）
- 标的公司审计报告或模拟报表
- 上市公司年报或审计报告

---

## 📥 Excel 模板下载

### DCF 建模模板

| 模板 | 场景 | 下载链接 |
|:---|:---|:---|
| 通用模板 | 基础 DCF 结构 | [下载](https://raw.githubusercontent.com/tedlinforlearning-question/a-share-corp-actions-value-impact/main/excel_templates/DCF_Model_Template.xlsx) |
| IPO 专用 | 募投项目 + 在建工程 | [下载](https://raw.githubusercontent.com/tedlinforlearning-question/a-share-corp-actions-value-impact/main/excel_templates/DCF_Model_Template_IPO.xlsx) |
| 并购买方 | 标的 FCFF + 发股/现金 + 业绩承诺 + 配套募资（条件启用，最多12 Sheet） | [下载](https://raw.githubusercontent.com/tedlinforlearning-question/a-share-corp-actions-value-impact/main/excel_templates/DCF_Model_Template_MA_Buy.xlsx) |
| 并购卖方 | 扣非分离 + 出售对比 | [下载](https://raw.githubusercontent.com/tedlinforlearning-question/a-share-corp-actions-value-impact/main/excel_templates/DCF_Model_Template_MA_Sale.xlsx) |

### 再融资分析模板

| 模板 | 场景 | Sheets | 下载链接 |
|:---|:---|:---:|:---|
| 纯补流摊薄 | 持股/BPS/EPS/利息节约/Pro Forma | 7 | [下载](https://raw.githubusercontent.com/tedlinforlearning-question/a-share-corp-actions-value-impact/main/excel_templates/Refinancing_WC_Dilution_Template.xlsx) |
| 补流+募投 | 摊薄 + 项目DCF + 合并估值 | 10 | [下载](https://raw.githubusercontent.com/tedlinforlearning-question/a-share-corp-actions-value-impact/main/excel_templates/Refinancing_WC_Project_DCF_Template.xlsx) |
| 收购+配套 | 摊薄 + 标的DCF + 备考合并 + 协同 | 7 | [下载](https://raw.githubusercontent.com/tedlinforlearning-question/a-share-corp-actions-value-impact/main/excel_templates/Refinancing_Acquisition_Template.xlsx) |

右键链接 → "另存为" 即可下载。

---

## 目录结构

```
A股IPO-重组-再融资-DCF估值/
├── SKILL.md                # 核心 Skill 配置（工作流定义）
├── README.md               # 本文件（项目说明）
├── ABOUT.md                # 项目背景与合作信息
├── docs/                   # 说明文档
│   ├── install_guide.md     # ⭐ 安装指南
│   ├── usage_guide.md      # 详细使用指南
│   └── limitations_and_roadmap.md # 局限性与后续计划
├── excel_templates/        # Excel 模板（📥 从 GitHub 下载）
├── scripts/                # 辅助脚本
│   ├── build_dcf_model.py      # DCF 模型自动构建
│   └── build_dilution.py       # 再融资摊薄分析
└── references/             # 参考资料
    └── case_params.md      # 案例参数速查（脱敏版）
```

---

## 关键特性

### 1. 募投项目作为"增量标的"

将传统 M&A 估值框架中的"标的公司"替换为"募投项目"，解决 IPO 场景下没有明确收购标的的估值难题。

### 2. 在建工程交叉比对

自动比对招股书中的募投项目与审计报告中的在建工程明细，修正招股书"T+4达产"的乐观假设，基于实际工程进度推演投产时间线。

### 3. 铁律：禁止硬编码

所有日期/年份自动从报告中提取。代码中变量名使用 `base_year` / `report_end_date` 等通用名，确保 Skill 适用于任意基准日的项目。

### 4. 备考扣非分离（出售特有）

对 M&A 出售项目的备考利润表进行扣非处理，避免大额处置收益/一次性投资收益污染经营性利润，确保 DCF 基于真实持续经营能力。

---

## 版本历史

| 版本 | 日期 | 更新内容 |
|:---|:---|:---|
| v1.3.0 | 2026-06 | 新增业绩承诺可行性分析（压力测试+增长驱动力拆解）；新增配套募资子表（双重摊薄+募投DCF+合并调整）；M&A购买流程9步→11步；Sheet结构8→12（条件扩展） |
| v1.2.0 | 2026-06 | 新增收购+配套融资场景（标的DCF+备考合并+协同量化）；六场景决策树完整 |
| v1.1.0 | 2026-06 | 新增再融资两大场景（纯补流摊薄 + 补流募投DCF）；五场景决策树；云端依赖改可选；脱敏 |
| v1.0.0 | 2026-05 | 初始版本，覆盖 IPO/M&A置入/M&A出售 三种场景 |

---

## 贡献

欢迎提交 Issue 和 Pull Request。

### 贡献方向

- 新增行业特定的 DCF 参数模板
- 补充更多可比公司 βU 数据源
- 完善 OCR 扫描版 PDF 支持
- 新增再融资/债券估值场景

---

## 关于作者

本 Skill 由投行从业者基于一线 IPO 与 M&A 项目经验构建。如果你也在做估值建模相关工作，欢迎交流。合作或工作机会请联系：_（在 GitHub Issue 中留言获取联系方式）_

---

## 许可

MIT License
