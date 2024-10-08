import io

from flask import Flask, flash, redirect, render_template, request, send_file

from scripts import converter
from utils import allowed_file

app = Flask(__name__,
            static_url_path='',
            template_folder='../../web/templates',
            static_folder='../../web/static')
app.secret_key = 'secret'


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/convert', methods=['POST'])
def convert():
    if 'file' not in request.files:
        flash('Nie wybrano pliku')
        return redirect('/')

    file = request.files.get('file', None)
    if not file or file.filename == '':
        flash('Nie wybrano pliku')
        return redirect('/')
    allowed = allowed_file(file.filename)
    if not allowed:
        flash("Nieprawidłowe rozszerzenie pliku. Wybierz plik z rozszerzeniem '.xls' lub '.xlsx'")
        return redirect('/')
    if file and allowed:
        try:
            converted = converter.convert(file)
        except ValueError:
            flash("Brak kolumny 'Data' i/lub 'Wydarzenie' w pliku Excel")
            return redirect('/')
        binary = io.BytesIO(converted.getvalue().encode())
        converted.close()
        return send_file(binary, as_attachment=True, download_name='output.ics')
    return redirect('/')
