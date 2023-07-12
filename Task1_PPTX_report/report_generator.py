import json
import csv
import os
from pptx import Presentation
from pptx.util import Inches
import matplotlib.pyplot as plt
import numpy as np

def readFile():
    with open("Task1_PPTX_report/sample.json", "r") as f:
        data = json.load(f)
    return data

def createPresentation(data):
    presentation = Presentation()

    for item in data['presentation']:
        if item['type'] == 'title':
            slide = presentation.slides.add_slide(presentation.slide_layouts[0])
            title = slide.shapes.title
            subtitle = slide.placeholders[1]

            title.text = item['title']
            subtitle.text = item['content']

        elif item['type'] == 'text':
            slide = presentation.slides.add_slide(presentation.slide_layouts[5])
            title = slide.shapes.title
            content = slide.shapes.add_textbox(Inches(1), Inches(1.5), Inches(1), Inches(1)).text_frame

            title.text = item['title']
            content.text = item['content']

        elif item['type'] == 'list':
            slide = presentation.slides.add_slide(presentation.slide_layouts[1])
            title = slide.shapes.title

            title.text = item['title']

            for thing in item['content']:
                if thing['level'] == 1:
                    content = slide.shapes.placeholders[1].text_frame.add_paragraph()
                    content.text = thing['text']
                    content.level = 1
                elif thing['level'] == 2:
                    content = slide.shapes.placeholders[1].text_frame.add_paragraph()
                    content.text = thing['text']
                    content.level = 2

            
        elif item['type'] == 'picture':
            slide = presentation.slides.add_slide(presentation.slide_layouts[5])
            title = slide.shapes.title
            pic = item['content']

            title.text = item['title']
            
            slide.shapes.add_picture(pic,Inches(3),Inches(2))

        elif item['type'] == 'plot':
            slide = presentation.slides.add_slide(presentation.slide_layouts[5])
            title = slide.shapes.title

            title.text = item['title']

            x = []
            y = []

            with open(item['content'], 'r') as f:
                csv_reader = csv.reader(f, delimiter=';')
            
                for row in csv_reader:
                    x.append(row[0])
                    y.append(row[1])

            # Create the XY plot using matplotlib
            plt.plot(x, y, 'b-')
            plt.xlabel(item['configuration']['x-label'])
            plt.ylabel(item['configuration']['y-label'])

            # Save the plot as an image
            plt.savefig('xy_plot.png')

            # Close the plot to free up resources
            plt.close()

            # Add the image to the slide
            x, y, cx, cy = Inches(2), Inches(2), Inches(6), Inches(4)
            slide.shapes.add_picture('xy_plot.png', x, y, cx, cy)

            


    return presentation


data = readFile()
presentation = createPresentation(data)
presentation.save("example.pptx")
