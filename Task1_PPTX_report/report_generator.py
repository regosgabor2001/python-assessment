import json
import csv
import os
from pptx import Presentation
from pptx.util import Inches
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import XySeriesData,XyChartData

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

            xValues = []
            yValues = []

            with open(item['content'], 'r') as f:
                csv_reader = csv.reader(f, delimiter=';')
            
                for row in csv_reader:
                    xValues.append(row[0])
                    yValues.append(row[1])

            chart_data = XyChartData()
            cd = chart_data.add_series('',number_format=None)

            for x, y in list(zip(xValues, yValues)):
                cd.add_data_point(x, y, number_format=None)

            x, y, cx, cy = Inches(1), Inches(2), Inches(6), Inches(6)
            chart = slide.shapes.add_chart(XL_CHART_TYPE.XY_SCATTER_LINES_NO_MARKERS, x, y, cx, cy, chart_data).chart
            chart.category_axis.axis_title.text_frame.text= "XTitle"
            chart.value_axis.axis_title.text_frame.text= "YTitle"

            #chart.chart_title.text_frame.text='ChartTitle'

            


    return presentation


data = readFile()
presentation = createPresentation(data)
presentation.save("example.pptx")
