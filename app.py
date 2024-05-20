from flask import Flask, request, render_template, send_file
import pandas as pd
import os

app = Flask(__name__)

# Helper function to convert numeric to boolean string
def numeric_to_bool_str(val):
    if val == 1.0:
        return 'TRUE'
    elif val == 0.0:
        return 'FALSE'
    return val

@app.route('/')
def upload_files():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'yaml_file' not in request.files or 'matrix_file' not in request.files:
        return "No file part"
    
    yaml_file = request.files['yaml_file']
    matrix_file = request.files['matrix_file']

    if yaml_file.filename == '' or matrix_file.filename == '':
        return "No selected file"

    yaml_df = pd.read_excel(yaml_file)
    matrix_df = pd.read_excel(matrix_file, sheet_name='Sheet1')

    # Normalize the column names for easier comparison
    yaml_df.columns = [col.strip() for col in yaml_df.columns]
    matrix_df.columns = [col.strip() for col in matrix_df.columns]

    # Normalize the configuration item names for easier comparison
    yaml_df['Configuration Item Name'] = yaml_df['Configuration Item Name'].str.strip().str.upper()
    matrix_df['Unnamed: 0'] = matrix_df['Unnamed: 0'].str.strip().str.upper()

    # Identify exporter-related columns
    exporter_columns = [col for col in yaml_df.columns if 'Exporter' in col]

    # Initialize a list to collect the report data with additional columns for IP Address and FQDN
    report = []

    # Iterate over each row in the YAML configuration
    for _, config_item in yaml_df.iterrows():
        config_name = config_item['Configuration Item Name']
        ip_address = config_item['IP Address']
        fqdn = config_item['FQDN']
        matching_exporter_row = matrix_df.loc[matrix_df['Unnamed: 0'] == config_name]

        if not matching_exporter_row.empty:
            for exporter in matching_exporter_row.columns[1:]:
                required_exporter = matching_exporter_row[exporter].values[0]
                if required_exporter == 'Y':
                    if exporter == 'exporter_blackbox':
                        icmp_val = numeric_to_bool_str(config_item.get('icmp'))
                        ssh_banner_val = numeric_to_bool_str(config_item.get('ssh-banner'))
                        tcp_connect_val = numeric_to_bool_str(config_item.get('tcp-connect'))
                        if not (icmp_val == 'TRUE' or ssh_banner_val == 'TRUE' or tcp_connect_val == 'TRUE'):
                            report.append({
                                'Configuration Item Name': config_name,
                                'IP Address': ip_address,
                                'FQDN': fqdn,
                                'Missing Exporter': 'blackbox'
                            })
                    elif exporter == 'exporter_ssl':
                        if config_item.get('Exporter_SSL') != 'TRUE':
                            report.append({
                                'Configuration Item Name': config_name,
                                'IP Address': ip_address,
                                'FQDN': fqdn,
                                'Missing Exporter': 'ssl'
                            })
                    else:
                        if not any(config_item.get(col) == exporter for col in exporter_columns):
                            report.append({
                                'Configuration Item Name': config_name,
                                'IP Address': ip_address,
                                'FQDN': fqdn,
                                'Missing Exporter': exporter
                            })

    # Convert the report to a DataFrame for easier handling
    report_df = pd.DataFrame(report)

    # Save the report to an Excel file
    report_path = 'exporter_report.xlsx'
    report_df.to_excel(report_path, index=False)

    return send_file(report_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
