import os
import pandas as pd
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
import re
import io
from werkzeug.utils import secure_filename
import logging

app = Flask(__name__)
app.secret_key = 'sua_chave_secreta_aqui'  # Necessário para flash messages

ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# Configuração de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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
            logger.error('Nenhum arquivo foi enviado na requisição.')
            return redirect(request.url)

        file = request.files['file']

        if file.filename == '':
            flash('Nenhum arquivo selecionado')
            logger.error('Nenhum arquivo foi selecionado para upload.')
            return redirect(request.url)

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            logger.info(f'Arquivo recebido: {filename}')

            # Processar o arquivo Excel diretamente na memória
            try:
                df = pd.read_excel(file, header=None)
                logger.info(f'Arquivo {filename} carregado com sucesso.')
            except Exception as e:
                flash(f'Erro ao ler o arquivo Excel: {e}')
                logger.exception('Erro ao ler o arquivo Excel.')
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
                            logger.info(f'Bloco finalizado: {current_activity} - {len(temp_df)} linhas capturadas.')

                        block_data = []  # Reiniciar a coleta de dados para o novo bloco
                        is_collecting = True

                    elif is_collecting and row[0] == 'Started By':  # Linha que define as colunas
                        logger.debug('Linha de cabeçalho "Started By" encontrada. Pulando.')
                        continue  # Pula a linha das colunas

                    elif is_collecting and 'Total' in str(row[0]):  # Finalizar o bloco ao encontrar "Total"
                        temp_df = pd.DataFrame(block_data, columns=columns)
                        all_results.append(temp_df)
                        logger.info(f'Bloco finalizado com "Total": {current_activity} - {len(temp_df)} linhas capturadas.')
                        block_data = []  # Reiniciar a coleta de dados para o próximo bloco

                    elif is_collecting:  # Coletar todas as linhas até encontrar "Total"
                        try:
                            start_date = row[2]  # Considerando que Start Date está na coluna 2
                            duration = row[6]  # Considerando que Duration está na coluna 6
                            decimal_duration = pd.to_timedelta(duration).total_seconds() / 3600  # Converter duração para decimal
                            block_data.append([current_activity, start_date, duration, round(decimal_duration, 2)])
                        except KeyError as e:
                            logger.error(f"Erro ao acessar coluna: {e}. Verifique se a coluna existe.")
                            continue

            # Após percorrer todas as linhas, certifique-se de que o último bloco seja capturado
            if block_data:
                temp_df = pd.DataFrame(block_data, columns=columns)
                all_results.append(temp_df)
                logger.info(f'Bloco finalizado: {current_activity} - {len(temp_df)} linhas capturadas.')

            # Combinar todos os blocos processados em um único DataFrame
            try:
                final_df = pd.concat(all_results, ignore_index=True)
                logger.info('Todos os blocos combinados em um único DataFrame.')
            except ValueError as e:
                flash(f'Erro ao combinar os dados processados: {e}')
                logger.exception('Erro ao combinar os DataFrames.')
                return redirect(url_for('index'))

            # Remover linhas vazias
            final_df.dropna(inplace=True)
            logger.info('Linhas vazias removidas do DataFrame final.')

            # Converter a coluna 'Start Date' para datetime
            try:
                final_df['Start Date'] = pd.to_datetime(final_df['Start Date'], errors='coerce')
                final_df.dropna(subset=['Start Date'], inplace=True)  # Remover linhas onde a conversão falhou
                logger.info('Coluna "Start Date" convertida para datetime.')
            except Exception as e:
                flash(f'Erro ao converter a coluna Start Date para datetime: {e}')
                logger.exception('Erro ao converter a coluna Start Date para datetime.')
                return redirect(url_for('index'))

            # Obter as datas de início e fim do formulário
            start_date_str = request.form.get('inicio_date')
            end_date_str = request.form.get('fim_date')

            if not start_date_str or not end_date_str:
                flash('Por favor, selecione ambas as datas.')
                logger.error('Datas de início ou fim não foram fornecidas.')
                return redirect(request.url)

            # Converter as datas selecionadas para datetime
            try:
                inicio_dt = pd.to_datetime(start_date_str)
                fim_dt = pd.to_datetime(end_date_str)
                logger.info(f'Datas selecionadas - Início: {inicio_dt}, Fim: {fim_dt}')
            except Exception as e:
                flash(f"Erro ao converter as datas: {e}")
                logger.exception('Erro ao converter as datas selecionadas.')
                return redirect(url_for('index'))

            # Filtrar pelo intervalo de datas selecionado
            try:
                filtered_df = final_df[(final_df['Start Date'] >= inicio_dt) & (final_df['Start Date'] <= fim_dt)]
                logger.info(f'Dados filtrados: {len(filtered_df)} linhas encontradas no intervalo.')
            except Exception as e:
                flash(f"Erro ao filtrar os dados: {e}")
                logger.exception('Erro ao filtrar os dados.')
                return redirect(url_for('index'))

            if filtered_df.empty:
                flash('Nenhum dado encontrado no intervalo de datas selecionado.')
                logger.warning('Nenhum dado encontrado após a filtragem.')
                return redirect(request.url)

            # Criar um arquivo Excel em memória
            output = io.BytesIO()
            try:
                filtered_df.to_excel(output, index=False)
                output.seek(0)  # Voltar para o início do BytesIO para leitura
                logger.info('Arquivo Excel filtrado criado em memória.')
            except Exception as e:
                flash(f'Erro ao salvar o arquivo filtrado: {e}')
                logger.exception('Erro ao salvar o arquivo filtrado em memória.')
                return redirect(request.url)

            return send_file(
                output,
                as_attachment=True,
                download_name='resultado_filtrado_por_datas.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        # Rota para método GET
        return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=True)
