# 安装指南

> 本指南帮助你在 15 分钟内完成所有前置准备，开始使用 IPO & M&A 估值 DCF 建模 Skill。

---

## 第一步：确认你有 Python

### 1.1 检查是否已安装

打开终端（Windows 按 `Win+R` 输入 `cmd`），输入：

```bash
python --version
```

如果显示 `Python 3.9.x` 或更高版本 → ✅ 跳到第二步  
如果显示 `'python' 不是内部或外部命令` → 需要安装

### 1.2 安装 Python（如需要）

1. 打开 Python 官网：https://www.python.org/downloads/
2. 点击黄色按钮下载最新版（3.11 或 3.12 均可）
3. 运行安装程序，**务必勾选** ☑ `Add Python to PATH`
4. 点击 Install Now，等待完成
5. 重新打开终端，再次输入 `python --version` 确认

---

## 第二步：安装 Python 依赖包

在终端中依次执行以下命令：

```bash
# 1. PDF 读取（用于提取年报/招股书中的财务数据）
pip install PyMuPDF

# 2. Excel 生成（用于输出估值模型 .xlsx 文件）
pip install openpyxl
```

这两个是**必须安装**的基础包。下面是可选的：

```bash
# 3. AKShare 金融数据（免费，如果没有 Wind 终端可以用这个）
pip install akshare

# 4. Tushare 金融数据（免费，备选方案）
pip install tushare
```

安装完成后验证：

```bash
pip list | grep -E "PyMuPDF|openpyxl|akshare"
```

应该看到 3 个包都已安装。

---

## 第三步：接入金融数据源（三选一）

> Skill 需要从金融数据源获取 β 值、可比公司财务、行业增速等数据。**至少接入一个**。

### 选项 A：Wind 万得终端（推荐，如果你公司有 Wind）

Wind 是投行最常用的金融数据终端，数据最全面。

#### A1. 获取 Wind API Key

1. 打开浏览器访问：**https://aifinmarket.wind.com.cn/#/docs**
2. 使用 Wind 账号登录（应该是公司分配给你的 Wind 账号）
3. 进入"API 管理"页面
4. 点击"创建 API Key"，复制生成的 Key（格式如 `WIND_xxxxxxxxxxxx`）
5. **这个 Key 只显示一次，务必复制保存好**

#### A2. 安装 Wind Skill

在 WorkBuddy/CodeBuddy 中执行：

```
安装 wind-mcp-skill
```

或手动：

1. 将 `wind-mcp-skill` 目录放在 `~/.workbuddy/skills/wind-mcp-skill/`
2. 在对话中输入：`配置 WIND_API_KEY = 你的key`

#### A3. 验证

```
帮我查一下 000001.SZ 的最新行情
```

如果能返回数据 → ✅ 配置成功。

---

### 选项 B：AKShare（免费，无需账户）

如果你没有 Wind 终端，AKShare 是完全免费的替代方案。

```bash
pip install akshare
```

**无需注册、无需 API Key**，安装即用。首次调用时会自动下载数据缓存。

> ⚠ 注意：AKShare 的 β 值数据不如 Wind 完整，建议至少先用 AKShare，后续有条件再接入 Wind。

---

### 选项 C：Tushare（免费，需要注册）

```bash
pip install tushare
```

1. 访问 https://tushare.pro/ 注册账号
2. 在"个人主页"获取 API Token
3. 使用前调用 `ts.set_token('你的token')`

---

## 第四步：安装 Skill 到 WorkBuddy

### 方式一：直接拷贝（推荐）

1. 找到本目录（`E:\工作考试\AI skill\` 或你解压的位置）
2. 复制整个 `ipo-ma-dcf` 文件夹
3. 粘贴到 `C:\Users\你的用户名\.workbuddy\skills\` 目录下
4. 重启 WorkBuddy 或重新加载 Skills

最终目录结构应该是：

```
C:\Users\你的用户名\.workbuddy\skills\ipo-ma-dcf\
├── SKILL.md
├── README.md
├── docs/
├── scripts/
├── excel_templates/
└── references/
```

### 方式二：通过界面导入

1. 在 WorkBuddy 左侧菜单找到"技能/Skill"管理
2. 点击"导入"
3. 选择本目录，确认

---

## 第五步：验证安装

在 WorkBuddy/CodeBuddy 对话中输入：

```
ipo-ma-dcf skill 是否已加载？
```

如果能看到工作流说明 → ✅ 安装成功，可以开始使用。

---

## 常见问题排查

### Q1: `pip install PyMuPDF` 报错？

**Windows 常见问题：**

```bash
# 如果提示缺少 Microsoft Visual C++ Build Tools
# 方案一：使用预编译包
pip install --only-binary=:all: PyMuPDF

# 方案二：安装较低版本
pip install PyMuPDF==1.23.8
```

### Q2: Wind API Key 在哪里找？

记住网址：**https://aifinmarket.wind.com.cn/#/docs**  
用 Wind 账号登录 → "API 管理" → "创建 API Key"

### Q3: 审计报告是扫描版 PDF，提取不到数据怎么办？

Skill 会自动检测扫描版。如果检测到，会提示你：
- 从招股书中提取财务数据（招股书通常含文本版报表）
- 或安装 OCR 工具：`pip install easyocr`（速度慢但效果好）

### Q4: Excel 模板文件在哪里？

模板已上传到 GitHub，直接下载：

| 模板 | 下载链接 |
|:---|:---|
| 通用模板 | https://raw.githubusercontent.com/tedlinforlearning-question/a-share-corp-actions-value-impact/main/DCF_Model_Template.xlsx |
| IPO 专用 | https://raw.githubusercontent.com/tedlinforlearning-question/a-share-corp-actions-value-impact/main/DCF_Model_Template_IPO.xlsx |
| 并购买方 | https://raw.githubusercontent.com/tedlinforlearning-question/a-share-corp-actions-value-impact/main/DCF_Model_Template_MA_Buy.xlsx |
| 并购卖方 | https://raw.githubusercontent.com/tedlinforlearning-question/a-share-corp-actions-value-impact/main/DCF_Model_Template_MA_Sale.xlsx |

右键链接 → "另存为" 即可下载。也可以直接用 Python 脚本本地生成：

```bash
python scripts/build_dcf_model.py --all
```

### Q5: 我可以不用 WorkBuddy，直接在 Python 里用吗？

可以。`scripts/` 目录下的 Python 脚本可以独立使用：

```bash
# 从 PDF 提取财务数据
python scripts/extract_financials.py 招股书.pdf output.json

# 生成 DCF Excel 模型
python scripts/build_dcf_model.py --scenario ipo
```

---

## 快速检查清单

完成安装后，确认以下每一项：

- [ ] Python 3.9+ 已安装
- [ ] `pip install PyMuPDF openpyxl` 完成
- [ ] 至少一个数据源已接入（Wind / AKShare / Tushare）
- [ ] Skill 目录已拷贝到 `~/.workbuddy/skills/ipo-ma-dcf/`
- [ ] 在 WorkBuddy 中测试 `ipo-ma-dcf` 可以正常加载

全部打勾 → 可以开始估值建模了 🎉
