# 组会汇报 PPT

[直接下载 deck.pptx](https://github.com/shuihongyan7-max/test-codex/raw/refs/heads/main/deck.pptx)

`deck.pptx`：38页（正文32页、附录6页），含逐页讲稿。

## 重新生成

安装 Python 依赖：`pip install -r requirements.txt`。
另需安装 Noto CJK 字体（Ubuntu：`fonts-noto-cjk`）。
运行 `python build_deck.py`，在仓库根目录生成 `deck.pptx` 和讲解备注；图片及公式输出至 `assets/`。

`assets/` 包含生成所需的图片素材及生成图表；`slide_notes.json` 和 `组会汇报_讲解备注.md` 为逐页讲稿。
原有《大组会》文件保留。

## 审稿意见 Word

[下载审稿意见.docx](https://github.com/shuihongyan7-max/test-codex/raw/refs/heads/main/%E5%AE%A1%E7%A8%BF%E6%84%8F%E8%A7%81.docx)

共54条，按“原句＋问题”编号并标注正文页码，排除符号类问题。运行 `python build_review.py` 可重新生成。
