from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash, Response
import pandas as pd
import os
import uuid
import re
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'sua_chave_secreta_aqui'  # Necessário para flash messages

# Pasta para salvar uploads temporários
UPLOAD_FOLDER = 'uploads'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# Função para verificar a extensão do arquivo
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Rota principal para upload
@app.route('/')
def index():
    return render_template('upload.html')

# Rota para processar o upload
@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('Nenhum arquivo foi enviado')
        return redirect(request.url)
    
    file = request.files['file']
    
    if file.filename == '':
        flash('Nenhum arquivo selecionado')
        return redirect(request.url)
    
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        unique_id = str(uuid.uuid4())
        saved_filename = f"{unique_id}_{filename}"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], saved_filename)
        file.save(file_path)
        
        # Processar o arquivo Excel
        try:
            df = pd.read_excel(file_path, header=None)
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
    
        # Salvar o DataFrame processado temporariamente
        processed_filename = f"{unique_id}_processed.csv"
        processed_path = os.path.join(app.config['UPLOAD_FOLDER'], processed_filename)
        final_df.to_csv(processed_path, index=False)
        
        return render_template('filter.html', unique_id=unique_id)
    else:
        flash('Tipo de arquivo não permitido. Por favor, envie um arquivo .xlsx ou .xls')
        return redirect(request.url)

# Rota para processar a filtragem
@app.route('/filter', methods=['POST'])
def filter_data():
    unique_id = request.form.get('unique_id')
    start_date_str = request.form.get('start_date')
    end_date_str = request.form.get('end_date')
    
    if not unique_id or not start_date_str or not end_date_str:
        flash('Por favor, preencha todos os campos.')
        return redirect(url_for('index'))
    
    processed_filename = f"{unique_id}_processed.csv"
    processed_path = os.path.join(app.config['UPLOAD_FOLDER'], processed_filename)
    
    if not os.path.exists(processed_path):
        flash('Arquivo processado não encontrado. Por favor, faça o upload novamente.')
        return redirect(url_for('index'))
    
    try:
        final_df = pd.read_csv(processed_path, parse_dates=['Start Date'])
    except Exception as e:
        flash(f'Erro ao ler os dados processados: {e}')
        return redirect(url_for('index'))
    
    # Converter as datas selecionadas para datetime
    try:
        inicio_str = pd.to_datetime(start_date_str)
        fim_str = pd.to_datetime(end_date_str)
    except Exception as e:
        flash(f'Erro ao converter as datas: {e}')
        return redirect(url_for('index'))
    
    # Filtrar pelo intervalo de datas selecionado
    filtered_df = final_df[(final_df['Start Date'] >= inicio_str) & (final_df['Start Date'] <= fim_str)]
    
    if filtered_df.empty:
        flash('Nenhum dado encontrado no intervalo de datas selecionado.')
        return redirect(url_for('index'))
    
    # Salvar o resultado em um novo arquivo Excel
    resultado_filename = f"{unique_id}_resultado_filtrado_por_datas.xlsx"
    resultado_path = os.path.join(app.config['UPLOAD_FOLDER'], resultado_filename)
    try:
        filtered_df.to_excel(resultado_path, index=False)
    except Exception as e:
        flash(f'Erro ao salvar o arquivo filtrado: {e}')
        return redirect(url_for('index'))
    
    return render_template('result.html', filename=resultado_filename)

# Rota para download do arquivo filtrado e excluir após o download
@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if not os.path.exists(file_path):
        flash('Arquivo não encontrado.')
        return redirect(url_for('index'))
    
    def generate():
        try:
            with open(file_path, 'rb') as f:
                while True:
                    data = f.read(4096)
                    if not data:
                        break
                    yield data
        finally:
            # Após o término da leitura, delete o arquivo
            try:
                os.remove(file_path)
                print(f'Arquivo {filename} excluído com sucesso.')
            except Exception as e:
                print(f'Erro ao excluir o arquivo {filename}: {e}')
    
    return Response(generate(),
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={'Content-Disposition': f'attachment; filename={filename}'})

if __name__ == '__main__':
    app.run(debug=True)
