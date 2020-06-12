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
**添加形状**
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
### 02 使用演示文稿
python-pptx支持创建新的演示文稿，同时也支持修改已有的演示文稿。
实际上它只支持修改现有的演示文稿，只不过如果你从一个没有幻灯片的演示文稿开始，
乍一看像是创建了一个新的演示文稿。

然而，演示文稿的外观很大程度上取决于删去幻灯片后剩下的部分，
特别是主题、幻灯片母版以及由幻灯片模板派生出的版式。
让我们用示例一步一步操作，从演示文稿可以做的两件事开始：打开它和保存它。

**打开一个演示文稿**

最简单的入门方法是打开一个新的演示文稿，无需指定要打开的文件。
```
from pptx import Presentation

prs = Presentation()
prs.save('test.pptx')
```
这将从内置的模板创建一个新的演示文稿，不做任何修改将其保存为文件“test.pptx”。
需要主意以下几点：

· 所谓的默认模板只是一个PowerPoint文件，不包含任何幻灯片，和安装好的python-pptx包存储在一起。
这相当于从全新安装的PowerPoint新建一个新的演示文稿（基于白底的模板的4x3长宽比的演示文稿），
除了它不会包含任何幻灯片，PowerPoint默认情况下会添加一张空白的幻灯片。

· 保存之前，你无需对它进行任何操作，如果你想确切地了解模板包含地内容，
只需要查看由此模板创建的“test.pptx”文件。

· 我们称其为模板，但它只是删去所有幻灯片的PowerPoint文件。真正的PowerPoint模板文件（.potx文件）
有些不同，以后可能会有更多关于它的内容，但是在使用python-pptx的时候不需要用到它们。

**真正地打开一个演示文稿**

如果你想控制最终的演示文稿，或者对现有的演示文稿进行修改，则需要用文件名打开：
```
prs = Presentation('existing-prs-file.pptx')
prs.save('new-file-name.pptx')
```
注意事项：

· 你可以用这种方法打开任何2007或者更新版本的PowerPoint文件，
PowerPoint2003或者更旧版本的.ppt文件不能打开。尽管你可能不能操作所有的内容，
文件里面已经存在的东西可以加载和保存。功能设置仍在构建中，所以你还不能添加修改一些内容（例如批注页面）。
但是如果演示文稿已经存在这些内容，python-pptx足够礼貌，可以把它们放在那里，也足够聪明，
无需了解它们是什么含义，可以将它们保存下来。

· 如果你用同一个文件名打开和保存文件，python-pptx会覆盖原始文件，你需要明确你的目的。

**打开一个“file-like”的演示文稿**

python-pptx可以从所谓的file-like对象打开一个演示文稿，也可以保存为一个file-like对象。
这将便于从网络连接或者数据库中获取源/目标演示文稿，并且不会（也不允许）与系统产生交互。
实际上这意味着你可以传递一个打开的文件或者StringIO/BytesIO流对象来打开或者保存演示文稿，如下所示：
```
f = open('foobar.pptx')
prs = Presentation(f)
f.close()

# or

with open('foobar.pptx') as f:
    source_stream = StringIO(f.read())
prs = Presentation(source_stream)
source_stream.close()
...
target_stream = StringIO()
prs.save(target_stream)
```