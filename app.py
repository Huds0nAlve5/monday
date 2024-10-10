from flask import Flask, render_template, request, redirect, url_for, flash, Response
import pandas as pd
import os
import uuid
import re
from datetime import datetime
from werkzeug.utils import secure_filename
import boto3
from botocore.exceptions import NoCredentialsError

app = Flask(__name__)
app.secret_key = 'sua_chave_secreta_aqui'  # Necessário para flash messages

# Configuração do cliente S3
s3_client = boto3.client(
    's3',
    aws_access_key_id=os.getenv('AWS_ACCESS_KEY_ID'),
    aws_secret_access_key=os.getenv('AWS_SECRET_ACCESS_KEY')
)

# Nome do bucket S3
AWS_BUCKET_NAME = os.getenv('AWS_BUCKET_NAME')

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

        # Salvar o arquivo no S3
        try:
            s3_client.upload_fileobj(file, AWS_BUCKET_NAME, saved_filename)
            print(f'Arquivo enviado: {saved_filename}')
        except NoCredentialsError:
            flash('Credenciais AWS não encontradas')
            return redirect(request.url)
        except Exception as e:
            flash(f'Erro ao enviar arquivo para S3: {e}')
            return redirect(request.url)

        # Processar o arquivo Excel diretamente da memória
        try:
            # Ler o arquivo Excel diretamente do S3
            response = s3_client.get_object(Bucket=AWS_BUCKET_NAME, Key=saved_filename)
            df = pd.read_excel(response['Body'], header=None)  # Carregar diretamente do objeto do S3
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
        columns = ['Activity', 'Start Date', 'Duration', 'Decimal']
        
        # Dicionário para armazenar a atividade e garantir consistência nos nomes
        activity_dict = {}
        
        pattern = r"[a-zA-Z]-\d"
        # Percorrer todas as linhas para processar os blocos
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
                        print(f"Erro ao acessar coluna: {e}. Verifique se a coluna existe.")
                        continue

        if block_data:
            temp_df = pd.DataFrame(block_data, columns=columns)
            all_results.append(temp_df)
            print(f'Bloco finalizado: {current_activity} - {len(temp_df)} linhas capturadas')
    
        final_df = pd.concat(all_results, ignore_index=True)
        final_df.dropna(inplace=True)
        
        # Salvar o DataFrame processado temporariamente no S3
        processed_filename = f"{unique_id}_processed.csv"
        processed_path = os.path.join(app.config['UPLOAD_FOLDER'], processed_filename)
        final_df.to_csv(processed_path, index=False)
        
        # Enviar o arquivo processado para o S3
        try:
            s3_client.upload_file(processed_path, AWS_BUCKET_NAME, processed_filename)
        except Exception as e:
            flash(f'Erro ao enviar arquivo processado para S3: {e}')
            return redirect(url_for('index'))
        
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
    
    try:
        response = s3_client.get_object(Bucket=AWS_BUCKET_NAME, Key=processed_filename)
        final_df = pd.read_csv(response['Body'], parse_dates=['Start Date'])
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
    try:
        filtered_df.to_excel(resultado_filename, index=False)
        # Enviar o arquivo filtrado para S3
        s3_client.upload_file(resultado_filename, AWS_BUCKET_NAME, resultado_filename)
    except Exception as e:
        flash(f'Erro ao salvar o arquivo filtrado: {e}')
        return redirect(url_for('index'))
    
    return render_template('result.html', filename=resultado_filename)

# Rota para download do arquivo filtrado e excluir após o download
@app.route('/download/<filename>')
def download_file(filename):
    try:
        response = s3_client.get_object(Bucket=AWS_BUCKET_NAME, Key=filename)
    except Exception as e:
        flash('Arquivo não encontrado.')
        return redirect(url_for('index'))
    
    def generate():
        try:
            yield response['Body'].read()
        finally:
            print(f'Arquivo {filename} lido com sucesso.')

    return Response(generate(),
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                    headers={'Content-Disposition': f'attachment; filename={filename}'})

if __name__ == '__main__':
    app.run(debug=True)
