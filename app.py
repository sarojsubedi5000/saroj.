from flask import Flask, render_template, request, redirect, url_for, send_file, session
import pandas as pd
import os
from nepali_datetime import date as nepali_date
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'your_secret_key'  # needed for session
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def detect_date_column(df):
    for column in df.columns:
        try:
            pd.to_datetime(df[column])
            return column
        except:
            continue
    return None

def convert_to_nepali_date(eng_date):
    try:
        date_obj = pd.to_datetime(eng_date).date()
        return nepali_date.from_datetime_date(date_obj)
    except:
        return ''

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    if file:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        session['filename'] = filename  # store filename in session
        return redirect(url_for('convert_page'))
    return "No file uploaded."

@app.route('/convert')
def convert_page():
    filename = session.get('filename')
    if not filename:
        return redirect(url_for('index'))
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    df = pd.read_excel(filepath)
    date_column = detect_date_column(df)
    return render_template('convert.html', filename=filename, date_column=date_column)

@app.route('/convert/execute', methods=['POST'])
def execute_conversion():
    filename = session.get('filename')
    if not filename:
        return redirect(url_for('index'))
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    df = pd.read_excel(filepath)
    date_column = detect_date_column(df)
    if date_column:
        df['Nepali Date'] = df[date_column].apply(convert_to_nepali_date)
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], f'converted_{filename}')
        df.to_excel(output_path, index=False)
        return send_file(output_path, as_attachment=True)
    return "No date column found."

if __name__ == '__main__':
    app.run(debug=True)
