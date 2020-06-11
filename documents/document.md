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

**Hello World!**
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
**带项目符号的幻灯片**
```
from pptx import Presentation

prs = Presentation()
bullet_slide_layout = prs.slide_layouts[1]

slide = prs.slides.add_slide(bullet_slide_layout)
shapes = slide.shapes

title_shape = shapes.title
body_shape = shapes.placeholders[1]

title_shape.text = 'Adding a Bullet Slide'

tf = body_shape.text_frame
tf.text = 'Find the bullet slide layout'

p = tf.add_paragraph()
p.text = 'Use _TextFrame.text for first bullet'
p.level = 1

p = tf.add_paragraph()
p.text = 'Use _TextFrame.add_paragraph() for subsequent bullets'
p.level = 2

prs.save('./result/example0102.pptx')
```
不是所有的形状都可以包含文本，但是那些可以包含文本的形状都至少包含一个段落，
即使是空段落或者形状里没有文本可见。
 _BaseShape.has_text_frame可用于确定一个形状能否包含文本。
 
 当_BaseShape.has_text_frame是
 True的时候，_BaseShape.text_frame.paragraphs[0]返回第一段，第一个段落的文本可以用
 text_frame.paragraphs[0].text设置。有一种捷径，可写属性_BaseShape.text 和 _TextFrame.text
 也可以实现相同的功能。后面两种方法在设置文本之前会删除形状内的所有文本，前面一种方法不会。
 
 **添加文本框**
```
from pptx import Presentation
from pptx.util import Inches, Pt

prs = Presentation()
blank_slide_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_slide_layout)

left = top = width = height = Inches(1)
txBox = slide.shapes.add_textbox(left, top, width, height)
tf = txBox.text_frame

tf.text = "This is text inside a textbox"

p = tf.add_paragraph()
p.text = "This is a second paragraph that's bold"
p.font.bold = True

p = tf.add_paragraph()
p.text = "This is a third paragraph that's big"
p.font.size = Pt(40)

prs.save('./result/example0103.pptx')
```
**添加图片**
```
from pptx import Presentation
from pptx.util import Inches

img_path = './image/monty-truth.png'

prs = Presentation()
blank_slide_layout = prs.slide_layouts[6]
slide = prs.slides.add_slide(blank_slide_layout)

left = top = Inches(1)
pic = slide.shapes.add_picture(img_path, left, top)

left = Inches(5)
height = Inches(5.5)
pic = slide.shapes.add_picture(img_path, left, top, height=height)

prs.save('./result/example0104.pptx')
```
添加形状
```
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

prs = Presentation()
title_only_slide_layout = prs.slide_layouts[5]
slide = prs.slides.add_slide(title_only_slide_layout)
shapes = slide.shapes

shapes.title.text = 'Adding an AutoShape'

left = Inches(0.93)  # 0.93" centers this overall set of shapes
top = Inches(3.0)
width = Inches(1.75)
height = Inches(1.0)

shape = shapes.add_shape(MSO_SHAPE.PENTAGON, left, top, width, height)
shape.text = 'Step 1'

left = left + width - Inches(0.4)
width = Inches(2.0)  # chevrons need more width for visual balance

for n in range(2, 6):
    shape = shapes.add_shape(MSO_SHAPE.CHEVRON, left, top, width, height)
    shape.text = 'Step %d' % n
    left = left + width - Inches(0.4)

prs.save('./result/example0105.pptx')
```
**添加表格**
```
from pptx import Presentation
from pptx.util import Inches

prs = Presentation()
title_only_slide_layout = prs.slide_layouts[5]
slide = prs.slides.add_slide(title_only_slide_layout)
shapes = slide.shapes

shapes.title.text = 'Adding a Table'

rows = cols = 2
left = top = Inches(2.0)
width = Inches(6.0)
height = Inches(0.8)

table = shapes.add_table(rows, cols, left, top, width, height).table

# set column widths
table.columns[0].width = Inches(2.0)
table.columns[1].width = Inches(4.0)

# write column headings
table.cell(0, 0).text = 'Foo'
table.cell(0, 1).text = 'Bar'

# write body cells
table.cell(1, 0).text = 'Baz'
table.cell(1, 1).text = 'Qux'

prs.save('./result/example0106.pptx')
```
**提取幻灯片中所有的文本**
```
from pptx import Presentation

prs = Presentation('./result/example0102.pptx')

# text_runs will be populated with a list of strings,
# one for each text run in presentation
text_runs = []

for slide in prs.slides:
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                text_runs.append(run.text)

print(text_runs)
```