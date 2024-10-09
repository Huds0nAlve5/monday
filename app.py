import os
import pandas as pd
from flask import Flask, render_template, request, send_file, flash, redirect, url_for, after_this_request
import re

app = Flask(__name__)
app.secret_key = 'sua_chave_secreta_aqui'
EXCEL_FILE_PATH = os.path.join(os.getcwd(), 'Arquivo.xlsx')  # Caminho do arquivo fixo

def processar_excel():
    df = pd.read_excel(EXCEL_FILE_PATH, header=None)
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
                    start_date = row[2]
                    duration = row[6]
                    decimal_duration = pd.to_timedelta(duration).total_seconds() / 3600
                    block_data.append([current_activity, start_date, duration, round(decimal_duration, 2)])
                except KeyError as e:
                    print(f"Erro ao acessar coluna: {e}")
                    continue

    if block_data:
        temp_df = pd.DataFrame(block_data, columns=columns)
        all_results.append(temp_df)

    final_df = pd.concat(all_results, ignore_index=True)
    final_df.dropna(inplace=True)

    return final_df

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/filtrar', methods=['POST'])
def filtrar():
    inicio_str = request.form['inicio_date']
    fim_str = request.form['fim_date']

    if not inicio_str or not fim_str:
        flash("Por favor, selecione ambas as datas.")
        return redirect(url_for('index'))

    final_df = processar_excel()
    final_df['Start Date'] = pd.to_datetime(final_df['Start Date'])

    try:
        inicio_dt = pd.to_datetime(inicio_str)
        fim_dt = pd.to_datetime(fim_str)
    except Exception as e:
        flash(f"Erro ao converter as datas: {e}")
        return redirect(url_for('index'))

    filtered_df = final_df[(final_df['Start Date'] >= inicio_dt) & (final_df['Start Date'] <= fim_dt)]

    if filtered_df.empty:
        flash("Nenhum dado encontrado no intervalo de datas selecionado.")
        return redirect(url_for('index'))

    filtered_file = 'resultado_filtrado_por_datas.xlsx'
    try:
        filtered_df.to_excel(filtered_file, index=False)
    except Exception as e:
        flash(f"Erro ao salvar o arquivo filtrado: {e}")
        return redirect(url_for('index'))

    response = send_file(filtered_file, as_attachment=True)
    return response

if __name__ == '__main__':
    app.run(debug=True)
