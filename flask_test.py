# app.py

from flask import Flask, request, render_template
from pptx import Presentation
from pptx.util import Inches
from os import listdir

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('flask_test.html')

@app.route('/convert', methods=['POST'])
def convert():
    directory = request.form['directory']
    filenames = [f for f in listdir(directory) if f.endswith('.jpg')]

    prs = Presentation()
    for filename in filenames:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        left = top = Inches(1)
        height = Inches(5)
        pic = slide.shapes.add_picture(directory + '/' + filename, left, top, height=height)

    prs.save('converted.pptx')
    return 'PPTX file has been converted.'

if __name__ == '__main__':
    app.run(debug=True)
