import json
import os
from pptx import Presentation
from pptx.util import Inches

def readFile():
    with open("Task1_PPTX_report/sample.json", "r") as f:
        data = json.load(f)
    return data


def create_presentation(data):
    presentation = Presentation()

    for item in data['presentation']:
        if item['type'] == 'title':
            slide = presentation.slides.add_slide(presentation.slide_layouts[0])
            title = slide.shapes.title
            subtitle = slide.placeholders[1]

            title.text = item['title']
            subtitle.text = item['content']

        elif item['type'] == 'text':
            slide = presentation.slides.add_slide(presentation.slide_layouts[1])
            title = slide.shapes.title
            content = slide.placeholders[1]

            title.text = item['title']
            content.text = item['content']

        elif item['type'] == 'list':
            slide = presentation.slides.add_slide(presentation.slide_layouts[1])
            title = slide.shapes.title
            content = slide.placeholders[1]

            title.text = item['title']
            #content.text = '\n'.join([f"{'*' * level} {text}" for level, text in item['content']])

        elif item['type'] == 'picture':
            slide = presentation.slides.add_slide(presentation.slide_layouts[5])
            title = slide.shapes.title
            pic = "Task1_PPTX_report/econ.png"

            title.text = item['title']
            
            slide.shapes.add_picture(pic,Inches(3),Inches(2))

        elif item['type'] == 'plot':
            slide = presentation.slides.add_slide(presentation.slide_layouts[1])
            title = slide.shapes.title
            content = slide.placeholders[1]

            title.text = item['title']
            content.text = f"Plot: {item['content']}"

    return presentation


data = readFile()
presentation = create_presentation(data)
presentation.save("example.pptx")
