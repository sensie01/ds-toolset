from flask import Flask, render_template, request
import os
import pandas as pd
import warnings

app = Flask(__name__)

# Specify the upload folder
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def calculate_total_working_time(file_path):
    try:
        # Read the Excel file into a DataFrame
        df = pd.read_excel(file_path, header=None, skiprows=5)

        # Check if the specified columns are present
        if len(df.columns) < 3:
            raise ValueError("Columns 'Last Occurred' and 'Cleared On' not found. Ensure the data starts from row 6.")

        # Extract the relevant columns
        last_occurred_column = df.iloc[:, 1]
        cleared_on_column = df.iloc[:, 2]

        # Suppress the datetime format inference warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            # Convert datetime strings to datetime objects
            last_occurred_column = pd.to_datetime(last_occurred_column, errors='coerce')
            cleared_on_column = pd.to_datetime(cleared_on_column, errors='coerce')

        # Calculate the working period
        working_period = cleared_on_column - last_occurred_column

        # Sum up all individual working periods
        total_working_time = working_period.sum()

        return f"{total_working_time}"
    except Exception as e:
        return f"An error occurred: {e}"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/calculate', methods=['POST'])
def calculate():
    try:
        if 'file' not in request.files:
            return render_template('index.html', result="Please select a file.")

        file = request.files['file']

        if file.filename == '':
            return render_template('index.html', result="Please select a file.")

        file_path = os.path.join(app.config['UPLOAD_FOLDER'], 'uploaded_file.xlsx')
        file.save(file_path)

        # Calculate total working time
        result = calculate_total_working_time(file_path)

        return render_template('index.html', result=result)
    except Exception as e:
        return render_template('index.html', result=f"An error occurred: {e}")

if __name__ == '__main__':
    os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
    app.run(debug=True)
