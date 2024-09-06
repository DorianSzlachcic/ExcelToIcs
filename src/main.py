from flask import Flask, Response, render_template, request, flash, send_file
from utils import allowed_file

app = Flask(__name__)


@app.route('/')
async def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
async def convert():
    if 'file' not in request.files:
        flash('No file part')
    file = request.files['file']

    if file.filename == '':
        flash('No selected file')
    if file and allowed_file(file.filename):
        converted = convert(file)
        return send_file(converted)
    return Response("404 Not Found", status=404)
