# docx-reader

读取 Word 文档时还原段落的自动编号。

## 背景

docx 把"编号格式"和"编号序号"分开存放：

- 编号格式（numFmt、lvlText、start、suff 等）集中定义在 `word/numbering.xml`。
- 段落自身只通过 `<w:numPr>` 记录 `numId + ilvl` 这一对引用。
- 真正显示的 `1.`、`一、`、`(A)` 是 Word 打开文档时按段落顺序边走边累加计数渲染出来的。

所以直接用 `python-docx` 读 `paragraph.text`，只能拿到段落正文，看不到自动编号。本项目把 `numbering.xml` 解析出来，自己维护一份计数，把编号文本还原后拼回段落前面。

## 使用

```python
from main import DocxReader

reader = DocxReader("test.docx")
for p in reader.paragraphs:
    print(p.text)
```

直接运行 `main.py` 可以对比"原始读取"和"还原后"的差异。

## 支持的编号格式

decimal / decimalZero、upperLetter / lowerLetter、upperRoman / lowerRoman、ordinal、cardinalText / ordinalText、chineseCounting、chineseLegalSimplified、ideographTraditional、ideographZodiac。
