# 经济学论文追踪系统 PaperTracker

> 输入关键词，自动抓取经济学顶刊最新论文，翻译摘要，导出 Excel，生成文献综述。

![Python](https://img.shields.io/badge/Python-3.10+-blue) ![PyQt6](https://img.shields.io/badge/GUI-PyQt6-green) ![License](https://img.shields.io/badge/license-MIT-lightgrey)

---

## 功能简介

- **关键词检索**：支持布尔运算符（AND `*`、OR `+`、NOT `-`、精确短语 `'...'`）
- **英文顶刊**：覆盖 22 本经济学顶刊，通过 CrossRef API 抓取，OpenAlex 补充被引次数
- **中文顶刊**：支持知网检索，覆盖经济/管理类核心期刊共 20 本
- **篇关摘过滤**：在标题、关键词、摘要中精确筛选，支持跨字段布尔匹配
- **AI 翻译**：批量翻译英文标题和摘要（Claude / DeepSeek 等兼容接口均可）
- **导出 Excel**：格式化表格，含标题（中/英）、摘要、被引次数、期刊、DOI 等字段
- **文献综述**：一键生成学术综述（中文/英文），超过 40 篇自动两阶段生成
- **知网 PDF 下载**：在嵌入式浏览器中登录知网后，可直接下载选中论文 PDF
- **Bug 反馈**：内置反馈入口，一键提交问题给开发者

---

## 覆盖期刊

### 英文期刊（22 本）

| 类别 | 期刊 |
|------|------|
| Top 5 | AER · QJE · JPE · REStud · Econometrica |
| 综合 | REStat · JEL · JEP · AEJ Applied · AEJ Policy · AEJ Macro · AEJ Micro |
| 金融 | JF · JFE · RFS · JFQA |
| 劳动/发展 | JOLE · JHR · JDE |
| 国际/计量 | JIE · JoE |
| 工作论文 | NBER Working Papers |

### 中文期刊（20 本，知网检索）

经济研究、经济学季刊、世界经济、管理世界、中国工业经济、金融研究、数量经济技术经济研究、经济科学、经济学报、中国农村经济、农业经济问题、财贸经济、统计研究、会计研究、审计研究、中国软科学、管理科学、系统工程理论与实践、南开管理评论、中国管理科学

---

## 快速开始

### 方式一：下载 exe（推荐，无需配置 Python）

前往 [Releases](../../releases) 页面下载最新版 `PaperTracker.exe`，双击运行。

### 方式二：源码运行

**环境要求**：Python 3.10+

```bash
pip install PyQt6 openpyxl
# 知网 PDF 下载功能还需要：
pip install PyQt6-WebEngine
```

```bash
python paper_tracker.py
```

---

## 使用说明

### 1. 填写关键词

在顶部输入框输入关键词，支持布尔运算符：

| 运算符 | 含义 | 示例 |
|--------|------|------|
| `*` | AND（且） | `minimum wage * employment` |
| `+` | OR（或） | `inequality + poverty` |
| `-` | NOT（非） | `trade * China - tariff` |
| `'...'` | 精确短语 | `'carbon tax'` |
| `()` | 优先分组 | `(FDI + trade) * growth` |

### 2. 配置 AI 接口（翻译 / 综述）

在右侧「模型设置」填入：
- **接口类型**：Anthropic 原生 或 OpenAI 兼容（支持 DeepSeek、豆包等）
- **API Key**：对应服务商的密钥
- **API 地址**：中转站地址（使用原生接口留空即可）

不填写 API Key 时，仍可正常检索和导出，仅翻译和综述功能不可用。

### 3. 知网登录（中文文献）

勾选「搜索中文文献（知网）」后，点击「知网登录」，在弹出的内嵌浏览器中完成校园网/机构认证，登录状态会自动保留。

### 4. 开始检索

点击「▶ 开始抓取」，程序自动：
1. 从 CrossRef / 知网抓取论文
2. 通过 OpenAlex 补充被引次数
3. 补全缺失摘要

抓取完成后可依次点击「翻译」→「导出 Excel」→「生成综述」。

---

## 常见问题

**Q：不配置 API Key 能用吗？**
可以，检索和导出功能正常，翻译和综述功能需要 API Key。

**Q：中文文献搜不到？**
需要在校园网或机构 VPN 环境下登录知网，且已勾选「搜索中文文献」。

**Q：被引次数显示为 0 或 -1？**
CrossRef 被引数据较少，程序会自动通过 OpenAlex 补充；部分较新论文确实尚无引用记录。

**Q：发现 Bug 怎么反馈？**
点击程序界面右上角的「🐛 反馈 Bug」按钮，填写描述后直接提交。

---

## 技术栈

- GUI：PyQt6
- 数据来源：CrossRef API · OpenAlex API · 知网（PyQt6-WebEngine）
- AI 接口：Anthropic Claude / OpenAI 兼容接口
- 导出：openpyxl
- 打包：PyInstaller

---

## License

MIT
