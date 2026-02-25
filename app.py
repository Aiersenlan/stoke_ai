import os
from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd

app = Flask(__name__)

# Ensure the template directory exists
os.makedirs('templates', exist_ok=True)
os.makedirs('static', exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/get_available_dates')
def get_available_dates():
    # Scan current directory for files matching market_analysis_YYYYMMDD.xlsx
    dates = []
    for filename in os.listdir('.'):
        if filename.startswith('market_analysis_') and filename.endswith('.xlsx'):
            date_str = filename.replace('market_analysis_', '').replace('.xlsx', '')
            if len(date_str) == 8:
                dates.append(date_str)
    dates.sort(reverse=True)
    return jsonify({'dates': dates})

@app.route('/get_report/<date_str>')
def get_report(date_str):
    filename = f'market_analysis_{date_str}.xlsx'
    if not os.path.exists(filename):
        return jsonify({'error': 'Report not found'}), 404
    
    try:
        # Read all sheets
        xls = pd.ExcelFile(filename)
        result = {}
        for sheet_name in xls.sheet_names:
            print(f"Reading sheet: {sheet_name}")
            # First, read without header to get the raw data layout
            df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
            
            # The structure is known: 
            # Row 0: Date
            # Row 1: Main headers (外資買超, etc)
            # Row 2: Sub headers
            # Row 3 onwards: Data
            
            # Extract header info
            main_headers = df.iloc[1].fillna('').tolist()
            sub_headers = df.iloc[2].fillna('').tolist()
            
            # Extract data
            data_rows = []
            for i in range(3, len(df)):
                row_data = df.iloc[i].fillna('').tolist()
                # Check if row is completely empty
                if not all(str(item).strip() == '' for item in row_data):
                    data_rows.append(row_data)
            
            result[sheet_name] = {
                'main_headers': main_headers,
                'sub_headers': sub_headers,
                'data': data_rows
            }
            
        return jsonify(result)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/download/<date_str>')
def download(date_str):
    filename = f'market_analysis_{date_str}.xlsx'
    if os.path.exists(filename):
        return send_file(filename, as_attachment=True)
    return 'File not found', 404

@app.route('/trigger_analysis', methods=['POST'])
def trigger_analysis():
    # Execute analyze.py to generate new data
    try:
        import subprocess
        
        target_date = None
        if request.is_json:
            data = request.json or {}
            target_date = data.get('date') # e.g. YYYY-MM-DD
            
        cmd = ["python", "analyze.py"]
        if target_date:
            target_date_str = target_date.replace("-", "")
            cmd.append(target_date_str)
            print(f"Manual trigger activated for date {target_date_str}...")
        else:
            print("Manual trigger activated for today...")
            
        subprocess.check_call(cmd)
        return jsonify({"status": "success", "message": "分析完成，新的報表已產出！"})
    except subprocess.CalledProcessError as e:
        return jsonify({"status": "error", "message": "分析失敗: 請確保該日期為交易日且有開盤資料。"}), 500
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"status": "error", "message": f"系統錯誤: {str(e)}"}), 500

if __name__ == '__main__':
    print("啟動網頁伺服器: http://localhost:5000")
    app.run(debug=True, port=5000)
