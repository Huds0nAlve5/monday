from flask import Flask, request, render_template, send_file
import pandas as pd
import re
import os

app = Flask(__name__)

@app.route('/')
def upload_file():
    return render_template('upload.html')

@app.route('/process', methods=['POST'])
def process_file():
    # Verificar se um arquivo foi enviado
    if 'file' not in request.files:
        return "Nenhum arquivo enviado."

    file = request.files['file']
    if file.filename == '':
        return "Nenhum arquivo selecionado."

    # Carregar o arquivo Excel
    df = pd.read_excel(file, header=None)
    all_results = []
    current_activity = None
    is_collecting = False
    columns = ['Activity', 'Start Date', 'Duration', 'Decimal']
    activity_dict = {}
    pattern = r"[a-zA-Z]-\d"
    block_data = []

    for index, row in df.iterrows():
        if pd.notnull(row[0]):
            if re.search(pattern, str(row[0])):
                activity_key = row[0].split()[0]
                current_activity = activity_dict.get(activity_key, row[0])
                activity_dict[activity_key] = current_activity

                if is_collecting and block_data:
                    temp_df = pd.DataFrame(block_data, columns=columns)
                    all_results.append(temp_df)

                block_data = []
                is_collecting = True

            elif is_collecting and row[0] == 'Started By':
                continue

            elif is_collecting and 'Total' in str(row[0]):
                temp_df = pd.DataFrame(block_data, columns=columns)
                all_results.append(temp_df)
                block_data = []

            elif is_collecting:
                try:
                    start_date = row[2]  # A data de início
                    duration = row[6]  # A duração
                    decimal_duration = pd.to_timedelta(duration).total_seconds() / 3600
                    block_data.append([current_activity, start_date, duration, round(decimal_duration, 2)])
                except KeyError as e:
                    continue

    if block_data:
        temp_df = pd.DataFrame(block_data, columns=columns)
        all_results.append(temp_df)

    final_df = pd.concat(all_results, ignore_index=True)
    final_df.dropna(inplace=True)

    # Obter datas do formulário
    inicio_date = request.form.get('inicio_date')
    fim_date = request.form.get('fim_date')

    if inicio_date and fim_date:
        # Filtrar os dados com base nas datas selecionadas
        final_df['Start Date'] = pd.to_datetime(final_df['Start Date'])
        inicio_str = pd.to_datetime(inicio_date)
        fim_str = pd.to_datetime(fim_date)
        filtered_df = final_df[(final_df['Start Date'] >= inicio_str) & (final_df['Start Date'] <= fim_str)]
    else:
        filtered_df = final_df  # Se não houver datas, não filtra

    # Salvar o resultado em um novo arquivo Excel
    output_path = 'resultado_filtrado_por_datas.xlsx'
    filtered_df.to_excel(output_path, index=False)

    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
