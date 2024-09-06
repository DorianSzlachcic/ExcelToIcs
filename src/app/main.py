import io
from flask import Flask, flash, redirect, render_template, request, send_file

from scripts import converter
from utils import allowed_file

app = Flask(__name__, template_folder='../../templates')
app.secret_key = 'secret'

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        flash('No file part')
        return redirect('/')
        
    file = request.files.get('file', None)
    if not file or file.filename == '':
        flash('No selected file')
        return redirect('/')
    if file and allowed_file(file.filename):
        converted = converter.convert(file)
        binary = io.BytesIO()
        binary.write(converted.getvalue().encode())
        binary.seek(0)
        return send_file(binary, as_attachment=True, download_name='output.ics')
    return redirect('/')
