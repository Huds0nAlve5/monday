import os
import pandas as pd
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import re
import io
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'sua_chave_secreta_aqui'  # Necessário para flash messages

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# Função para verificar a extensão do arquivo
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Rota principal para upload e filtragem
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Verificar se o arquivo está na requisição
        if 'file' not in request.files:
            flash('Nenhum arquivo foi enviado')
            return redirect(request.url)

        file = request.files['file']

        if file.filename == '':
            flash('Nenhum arquivo selecionado')
            return redirect(request.url)

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)

            # Processar o arquivo Excel diretamente na memória
            try:
                df = pd.read_excel(file, header=None)
                print(f'Arquivo carregado: {filename}')
            except Exception as e:
                flash(f'Erro ao ler o arquivo Excel: {e}')
                return redirect(url_for('index'))

            # Inicializar variáveis para armazenar os resultados
            all_results = []

            # Variáveis de controle para identificar os blocos
            current_activity = None
            is_collecting = False
            block_data = []
            columns = ['Activity', 'Start Date', 'Duration', 'Decimal']  # Adicionar coluna Decimal

            # Dicionário para armazenar a atividade e garantir consistência nos nomes
            activity_dict = {}

            pattern = r"[a-zA-Z]-\d"
            # Percorrer todas as linhas para processar os blocos
            for index, row in df.iterrows():
                if pd.notnull(row[0]):
                    if re.search(pattern, str(row[0])):  # Se for o início de um novo bloco (nome da atividade)
                        activity_key = row[0].split()[0]  # Identificar a chave da atividade (ex: EC-6103)
                        current_activity = activity_dict.get(activity_key, row[0])  # Usar o nome já registrado ou o atual
                        activity_dict[activity_key] = current_activity  # Garantir consistência futura para a mesma atividade

                        if is_collecting and block_data:
                            # Finalizar o bloco anterior se houver
                            temp_df = pd.DataFrame(block_data, columns=columns)
                            all_results.append(temp_df)

                        block_data = []  # Reiniciar a coleta de dados para o novo bloco
                        is_collecting = True

                    elif is_collecting and row[0] == 'Started By':  # Linha que define as colunas
                        continue  # Pula a linha das colunas

                    elif is_collecting and 'Total' in str(row[0]):  # Finalizar o bloco ao encontrar "Total"
                        temp_df = pd.DataFrame(block_data, columns=columns)
                        all_results.append(temp_df)
                        block_data = []  # Reiniciar a coleta de dados para o próximo bloco

                    elif is_collecting:  # Coletar todas as linhas até encontrar "Total"
                        try:
                            start_date = row[2]  # Considerando que Start Date está na coluna 2
                            duration = row[6]  # Considerando que Duration está na coluna 6
                            decimal_duration = pd.to_timedelta(duration).total_seconds() / 3600  # Converter duração para decimal
                            block_data.append([current_activity, start_date, duration, round(decimal_duration, 2)])
                        except KeyError as e:
                            print(f"Erro ao acessar coluna: {e}. Verifique se a coluna existe.")
                            continue

            # Após percorrer todas as linhas, certifique-se de que o último bloco seja capturado
            if block_data:
                temp_df = pd.DataFrame(block_data, columns=columns)
                all_results.append(temp_df)
                print(f'Bloco finalizado: {current_activity} - {len(temp_df)} linhas capturadas')

            # Combinar todos os blocos processados em um único DataFrame
            final_df = pd.concat(all_results, ignore_index=True)

            # Remover linhas vazias
            final_df.dropna(inplace=True)

            # Converter a coluna 'Start Date' para datetime
            try:
                final_df['Start Date'] = pd.to_datetime(final_df['Start Date'], errors='coerce')
                final_df.dropna(subset=['Start Date'], inplace=True)  # Remover linhas onde a conversão falhou
            except Exception as e:
                flash(f'Erro ao converter a coluna Start Date para datetime: {e}')
                return redirect(url_for('index'))

            # Obter as datas de início e fim do formulário
            start_date_str = request.form.get('inicio_date')
            end_date_str = request.form.get('fim_date')

            if not start_date_str or not end_date_str:
                flash('Por favor, selecione ambas as datas.')
                return redirect(request.url)

            # Converter as datas selecionadas para datetime
            try:
                inicio_dt = pd.to_datetime(start_date_str)
                fim_dt = pd.to_datetime(end_date_str)
            except Exception as e:
                flash(f"Erro ao converter as datas: {e}")
                return redirect(request.url)

            # Filtrar pelo intervalo de datas selecionado
            filtered_df = final_df[(final_df['Start Date'] >= inicio_dt) & (final_df['Start Date'] <= fim_dt)]

            if filtered_df.empty:
                flash('Nenhum dado encontrado no intervalo de datas selecionado.')
                return redirect(request.url)

            # Criar um arquivo Excel em memória
            output = io.BytesIO()
            try:
                filtered_df.to_excel(output, index=False)
                output.seek(0)  # Voltar para o início do BytesIO para leitura
            except Exception as e:
                flash(f'Erro ao salvar o arquivo filtrado: {e}')
                return redirect(request.url)

            return send_file(
                output,
                as_attachment=True,
                download_name='resultado_filtrado_por_datas.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        else:
            flash('Tipo de arquivo não permitido. Por favor, envie um arquivo .xlsx ou .xls')
            return redirect(request.url)

    return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=True)
