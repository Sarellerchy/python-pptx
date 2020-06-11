## python-pptx中文文档
python-pptx是一个用来创建和更新PowerPoint (.pptx) 文件的python库。

通常的用途是根据数据库中的内容生成自定义的PowerPoint演示文稿，在web应用中点击链接即可下载。
一些开发者用它来基于工作管理系统中保存的信息自动生成用于展示的工程状态报告。
它也可以用来对演示文稿库批量更新，或者仅仅用来自动生成一两张幻灯片，这些如果手动更新的话会很繁琐。

### 安装
python-pptx托管在PyPI上，可以简单地用pip来安装：

`pip install python-pptx
`

python-pptx依赖lxml包和Pillow包（python图像库PIL的现代版），图表功能依赖XlsxWriter库。
pip 和 easy_install都会为你装好这些依赖的包，如果用setup.py来安装的话，则需要自己安装这些依赖的库。

### 依赖条件
· Python 2.6, 2.7, 3.3, 3.4, 3.6

· lxml

· Pillow

· XlsxWriter

### 01 快速入门
尝试以下示例，了解怎样使用python-pptx。

Hello World!
```
from pptx import Presentation

prs = Presentation()
title_slide_layout = prs.slide_layouts[0]
slide = prs.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]

title.text = "Hello, World!"
subtitle.text = "python-pptx was here!"

prs.save('./result/example0101.pptx')
```


