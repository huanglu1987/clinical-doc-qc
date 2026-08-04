# FDA（美国）申报资料格式规范

> **出处**：
> - FDA, *Portable Document Format (PDF) Specifications* / *PDF Specifications for FDA Regulatory Submissions*
> - FDA, *Guidance for Industry: Providing Regulatory Submissions in Electronic Format*
> - ICH, *Specification for Submission Formats for eCTD*
>
> **核对日期**：2026-08-04。
> **注意**：FDA 文档规格文件会随 eCTD 版本更新，正式递交前请以 FDA 官网现行版本复核。

---

## 一、工具会自动判定的项

已固化在 `scripts/docqc.py` 的 `FDA_SPEC`，运行 `docqc.py check <文件> --profile fda` 判定。

### .docx 源文件

| 项目 | 规定 | 判定 |
|---|---|---|
| 字体 | **Times New Roman 或 Arial** | 主字体不符 → 错误；混用其他 → 警告 |
| 正文字号 | **12 pt**（叙述性正文） | 不符 → 警告；低于 9pt → 错误 |
| 表格字号 | **≥10 pt**（9pt 为可接受下限，更小应避免） | <9pt → 错误；9–10pt → 警告 |
| 页边距 | 四边 **≥1 英寸**；硬性下限：左 0.75in、其他 0.375in | 低于硬性下限 → 错误；低于 1in → 警告 |

> **页边距的两个数值来源不同**，本工具同时采用：
> - "四边至少 1 英寸"出自 *Providing Regulatory Submissions in Electronic Format*，作为**建议值**（警告线）
> - "左侧 3/4 英寸、其余 3/8 英寸"出自 *PDF Specifications*，作为**硬性下限**（错误线）
>
> 之所以保留两条线，是因为前者是留装订余量的通行做法，后者才是会被拒收的红线。

### .pdf 递交件

| 项目 | 规定 | 判定 |
|---|---|---|
| PDF 版本 | **1.4 – 1.7** | 超出 → 错误 |
| 安全设置 | 不得加密/口令保护（指定 FDA 表格除外） | 有加密 → 错误 |
| 动态内容 | **禁止 JavaScript** 与 OpenAction 脚本 | 检出 → 错误 |
| 可搜索性 | 文本必须可搜索，扫描件须 OCR | 提取不到文本 → 错误 |
| 书签 | ≥5 页的文档必须有书签；层级建议 **≤4 层** | 无书签 → 错误；>4 层 → 警告 |
| 字体嵌入 | 非标准字体**必须嵌入** | 未嵌入 → 错误 |
| 网页浏览 | 应设为 Fast Web View | 工具暂不判定，需人工在 Acrobat 中确认 |

---

## 二、工具不判、需人工确认的项

- **超链接**：应使用相对路径而非绝对路径，缩放设为 "Inherit Zoom"
- **目录**：5 页以上文档应包含目录
- **OCR 准确性**：工具只能判断"有没有文字层"，判断不了 OCR 认错字
- **交叉引用指向是否正确**：工具能查出引用的表号不存在，但查不出"引用了表 3、实际应该引用表 4"
- **Fast Web View**：需在 Acrobat 的文档属性中查看

---

## 三、与 CDE 的主要差异

见 [`cde-format-spec.md`](cde-format-spec.md) 第五节。

最容易出错的两点：

1. **纸张**：FDA 是 US Letter，CDE 是 A4。同一份 Word 直接换语言递交，纸张尺寸不会自动改。
2. **字体**：FDA 允许 Arial，CDE 只认 Times New Roman（西文）+ 宋体（中文）。
   把 CDE 版直接翻译成英文递 FDA 时字体通常没问题，反过来（FDA 版 Arial → 中文版）就会不合规。
