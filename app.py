from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import os
import csv
import json
from pdf_converter import convertpdf
from amc_processor import process_amc_files, fetch_nav, fetch_ebal, historical_nav
from dotenv import load_dotenv
load_dotenv()

pwd = os.getenv("STORAGE_DIR")
if not pwd:
    print(f"STORAGE_DIR not found in environment, using current directory")
    pwd = os.getcwd()
    file_path = os.path.join(pwd, '.env')
    with open(file_path, 'w', newline='') as f:
            f.write('STORAGE_DIR=' + pwd)
    print(f"Created .env file with STORAGE_DIR={pwd}")
    


app = Flask(__name__, static_folder="./dist", static_url_path="/")
CORS(app)
#CORS(app, origins="http://localhost:5173")

@app.route("/")
def serve_react():
    return send_from_directory(app.static_folder, "index.html")

@app.route('/convert', methods=['POST'])
def convert_pdf_to_csv():
    print(request.form)   # Debug: log form data (including password)
    
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    file = request.files['file']
    password = request.form.get('password')
    output_filename = request.form.get('outputFilename')
    userid = request.form.get('userid')
    UPLOAD_FOLDER = os.path.join(pwd,userid,'uploads')
    app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    
    if not output_filename:
        return jsonify({'error': 'Output filename is required'}), 400

    # print(f"Output filename: {output_filename}")
    folder = os.path.dirname(output_filename)
    folder = os.path.join(userid,folder)
    output_filename = os.path.join(pwd,userid, output_filename)
    # print(f"Output folder: {folder}")

    try:
        if not os.path.exists(folder):
            os.makedirs(folder)

        if file.filename == '':
            return jsonify({'error': 'No selected file'}), 400
        
        if not password:
            return jsonify({'error': 'Password is required'}), 400

        # Save the uploaded PDF file.
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(file_path)

        # Convert the PDF to CSV.
        csv_data = convertpdf(file_path, password, userid)
        if csv_data is None:
            print(f"Failed to convert PDF: {file_path}, csv data is received as None")
            return jsonify({'error': 'Failed to convert PDF. Check password and file.'}), 500

        # Write the full CSV data to the main output file.
        # print(f"Writing CSV data to: {output_filename}")
        with open(output_filename, 'w', newline='') as f:
            f.write(csv_data)

        # Process the CSV data to create separate AMC files,
        # errorlist.csv, and AMClist.csv.
        process_amc_files(csv_data, folder)

        # Return a successful response without AMC mapping.
        return jsonify({
            'filename': file.filename.replace(".pdf", ".csv"),
            'csvData': csv_data
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/fetchdata', methods=['GET'])
def fetch_data():
    output_filename = request.args.get('filename')
    if not output_filename:
        return jsonify({'error': 'Filename is required'}), 400

    # Determine the full path of the CSV file.
    output_path = os.path.join(pwd, output_filename)

    try:
        with open(output_path, 'r') as f:
            reader = csv.DictReader(f)
            data = list(reader)
        return jsonify({'data': data})
    except FileNotFoundError:
        return jsonify({'error': 'File not found'}), 404
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/listAMC', methods=['GET'])
def list_amc():
    """
    This route returns the AMC list stored in AMClist.csv as JSON.
    It expects an 'outputFilename' query parameter to locate the folder.
    """
    output_filename = request.args.get('outputFilename')
    if not output_filename:
        return jsonify({'error': 'Output filename parameter is required'}), 400

    folder = os.path.dirname(output_filename)
    amc_list_file = os.path.join(folder, "AMClist.csv")
    # print(f"AMC list file: {amc_list_file} in folder: {folder}, received output filename: {output_filename}")
    # process_amc_files(output_filename, folder)
    try:
        with open(amc_list_file, 'r') as f:
            reader = csv.DictReader(f)
            amc_list = list(reader)
        return jsonify({'amcFiles': amc_list})
    except Exception as e:
        print(f"Error reading AMC list: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/schemeTransactions', methods=['GET'])
def scheme_transactions():
    # print(f"Schem Transactions called with args: {request.args}")
    output_filename = request.args.get('outputFilename')
    scheme_name = request.args.get('scheme')
    if not output_filename or not scheme_name:
        return jsonify({'error': 'Missing parameters'}), 400
    folder = os.path.dirname(output_filename)
    csv_path = os.path.join(folder, os.path.basename(output_filename))  # or the main CSV
    # print(f"CSV path: {csv_path} and folder: {folder}")
    try:
        with open(csv_path, 'r') as f:
            reader = csv.DictReader(f)
            all_rows = list(reader)
        # print(f"Total rows in CSV: {len(all_rows)}")
        # Filter rows by scheme
        filtered = [row for row in all_rows if row.get('scheme', '').strip() == scheme_name.strip()]
        # print(f"fetching NAV for scheme: {scheme_name} with ISIN: {filtered[0].get('isin', '')}")
        nav, navdate = fetch_nav(filtered[0].get('isin', ''))  # Fetch NAV and NAV date
        # print(f"fetching eBalance for scheme: {scheme_name}, with ISIN: {filtered[0].get('isin', '')} and folder {folder}")
        effective_bal, cagr = fetch_ebal(filtered[0].get('isin'), folder)  # Fetch eBalance
        return jsonify({
            'transactions': filtered,
            'nav': nav,
            'nav_date': navdate,
            'effective_balance': effective_bal,
            'cagr': cagr,
            'amfi': filtered[0].get('amfi', '')
        })
    except Exception as e:
        print(f"Error fetching scheme transactions: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/refreshAMC', methods=['GET'])
def refresh_amc():
    try:
        userid = request.args.get('userid')
        password = request.args.get('password')
        input_filename = request.args.get('input')
        outputfile = request.args.get('output')
        if not input_filename:
            return jsonify({'error': 'Input filename parameter is required'}), 400
        #remove NAVAll.txt if present
        # if os.path.exists('NAVAll.txt'):
        #     os.remove('NAVAll.txt')
        csv_data = convertpdf(input_filename, password, userid)
        folder = os.path.dirname(outputfile)
        # print(f"Writing CSV data to: {outputfile}")
        with open(outputfile, 'w', newline='') as f:
            f.write(csv_data)
        process_amc_files(csv_data, folder)
        return jsonify({'status': 'Success'}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/historicalNAV', methods=['GET'])
def historical_nav_data():
    amfi = request.args.get('amfi')
    # print(f"Fetching historical NAV for AMFI: {amfi}")
    if not amfi:
        return jsonify({'error': 'AMFI parameter is required'}), 400
    try:
        return(historical_nav(amfi))
    except Exception as e:
        return jsonify({'error received': str(e)}), 500    

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=4000)
