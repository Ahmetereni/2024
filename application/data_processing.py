from flask import Flask, jsonify, request, render_template, send_file, flash
import pandas as pd
import numpy as np
import re
import io

def create_app():
    app = Flask(__name__)
    app.secret_key = 'your_secret_key_here'  # Replace with your actual secret key

    def clean_excel(data):
        for column in data.columns:
            data[column] = data[column].apply(lambda x: clean_cell_value(x) if pd.notnull(x) else x)
        return data

    def clean_cell_value(cell_value):
        if isinstance(cell_value, str):
            if re.search(r'\d', cell_value) and re.search(r'\D', cell_value):
                numbers = re.findall(r'\d+', cell_value)
                return ' '.join(numbers)
        return cell_value

    def group_and_average(data: pd.DataFrame) -> pd.DataFrame:
        numeric_data = data.apply(pd.to_numeric, errors='coerce')
        numeric_data = numeric_data.dropna(axis=1, how='all')
        grouped_data = numeric_data.groupby(data.iloc[:, 0]).mean()
        return grouped_data

    @app.route('/')
    def index():
        return render_template('cleaner.html')

    @app.route('/process-data', methods=['POST'])
    def process_data():
        if 'file' not in request.files or 'max_criteria' not in request.form:
            return jsonify({'error': 'No file part or max_criteria'}), 400

        file = request.files['file']
        max_criteria_input = request.form['max_criteria']

        if file.filename == '':
            return jsonify({'error': 'No selected file'}), 400

        if file and max_criteria_input:
            data = pd.read_excel(file)
            cleaned_data = clean_excel(data)
            grouped_data = group_and_average(cleaned_data)

            try:
                max_criteria = [int(x) - 1 for x in max_criteria_input.split(',')]
            except ValueError:
                return jsonify({'error': 'Invalid max_criteria format. Please enter comma-separated integers.'})

            index_names = grouped_data.index
            initial_data = grouped_data.copy()
            data = grouped_data

            normalized_data = pd.DataFrame()
            for i in range(data.shape[1]):
                if i in max_criteria:
                    normalized_data[data.columns[i]] = (data.iloc[:, i] - data.iloc[:, i].min()) / (data.iloc[:, i].max() - data.iloc[:, i].min())
                else:
                    normalized_data[data.columns[i]] = (data.iloc[:, i].max() - data.iloc[:, i]) / (data.iloc[:, i].max() - data.iloc[:, i].min())

            normalized_data.index = index_names

            std_dev = normalized_data.std()
            correlation_matrix = normalized_data.corr()
            contradiction_matrix = 1 - correlation_matrix
            sumCol = contradiction_matrix.sum(axis=1)
            cj = std_dev * sumCol
            wj = cj / cj.sum()

            try:
                criteria_matrix = initial_data.astype(float)
                weights = wj.values

                norm_matrix = criteria_matrix / np.sqrt((criteria_matrix**2).sum(axis=0))
                weighted_norm_matrix = norm_matrix * weights
                ideal_solution = weighted_norm_matrix.max(axis=0)
                negative_ideal_solution = weighted_norm_matrix.min(axis=0)

                separation_ideal = np.sqrt(((weighted_norm_matrix - ideal_solution)**2).sum(axis=1))
                separation_negative_ideal = np.sqrt(((weighted_norm_matrix - negative_ideal_solution)**2).sum(axis=1))

                relative_closeness = separation_negative_ideal / (separation_ideal + separation_negative_ideal)
                ranking = relative_closeness.argsort()[::-1] + 1

                results_df = pd.DataFrame(index=initial_data.index)
                results_df['Relative Closeness'] = relative_closeness
                results_df['Rank'] = ranking

                topsis_output = io.BytesIO()
                with pd.ExcelWriter(topsis_output, engine='xlsxwriter') as writer:
                    results_df.to_excel(writer, sheet_name='TOPSIS Results')

                topsis_output.seek(0)
                return send_file(topsis_output, as_attachment=True, download_name="Topsis.xlsx", mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

            except Exception as e:
                print(f"An error occurred: {e}")
                flash(f"An error occurred: {e}", "danger")
                return render_template('cleaner.html')

    return app

if __name__ == '__main__':
    app = create_app()
    app.run(debug=True)
