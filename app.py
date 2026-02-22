from flask import Flask, render_template, request, send_file
import os
import zipfile
from processador_excel import processar_excel
from gerador_word import gerar_word

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'

# Criar pastas se não existirem
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/gerar', methods=['POST'])
def gerar():
    arquivos = request.files.getlist('arquivo[]')

    if not arquivos or arquivos[0].filename == '':
        return "Nenhum arquivo enviado"

    caminhos_word = []

    for arquivo in arquivos:
        # Salvar o XLSX temporariamente
        caminho_xlsx = os.path.join(app.config['UPLOAD_FOLDER'], arquivo.filename)
        arquivo.save(caminho_xlsx)

        # Extrair dados
        dados = processar_excel(caminho_xlsx)

        # Gerar Word
        caminho_word = gerar_word(dados, app.config['OUTPUT_FOLDER'])
        caminhos_word.append(caminho_word)

    # Criar ZIP final
    zip_path = os.path.join(app.config['OUTPUT_FOLDER'], "relatorios.zip")
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for caminho in caminhos_word:
            zipf.write(caminho, os.path.basename(caminho))

    return send_file(zip_path, as_attachment=True, download_name="relatorios.zip")

if __name__ == '__main__':
    app.run()