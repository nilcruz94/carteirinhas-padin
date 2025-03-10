from flask import Flask, request, redirect, url_for, render_template_string, jsonify, session, flash
import pandas as pd
import os
import qrcode
from io import BytesIO
from werkzeug.utils import secure_filename
from functools import wraps

app = Flask(__name__)
app.secret_key = 'sua_chave_secreta'  # Altere para uma chave segura
ACCESS_TOKEN = "minha_senha"  # Token de acesso

app.config['UPLOAD_FOLDER'] = 'uploads'

# Lista de extensões permitidas
ALLOWED_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif'}

# Cria os diretórios necessários, se não existirem
if not os.path.exists('static/qr_codes'):
    os.makedirs('static/qr_codes')
if not os.path.exists('static/fotos'):
    os.makedirs('static/fotos')
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

def allowed_file(filename):
    _, ext = os.path.splitext(filename)
    return ext.lower() in ALLOWED_EXTENSIONS

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login", next=request.url))
        return f(*args, **kwargs)
    return decorated_function

def gerar_html_carteirinhas(arquivo_excel):
    planilha = pd.read_excel(arquivo_excel, sheet_name='LISTA CORRIDA')
    dados = planilha[['RM', 'NOME', 'DATA NASC.', 'RA', 'SAI SOZINHO?', 'SÉRIE', 'HORÁRIO']]
    dados['RM'] = dados['RM'].fillna(0).astype(int)
    
    # Nesta página (geração das carteirinhas) removemos header e footer para não interferir na impressão.
    # Adicionamos também um overlay de loading e uma mensagem de sucesso para indicar que o processo foi concluído.
    html_content = """
<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Carteirinhas</title>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
  <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
  <style>
    body {
      font-family: 'Montserrat', sans-serif;
      margin: 0;
      padding: 20px;
      background-color: #f4f4f4;
      display: flex;
      flex-direction: column;
      align-items: center;
    }
    #search-container {
      margin-bottom: 20px;
    }
    #localizarAluno {
      padding: 8px;
      font-size: 16px;
      width: 300px;
    }
    .carteirinhas-container {
      width: 100%;
      max-width: 1100px;
    }
    .page {
      margin-bottom: 40px;
      position: relative;
    }
    .page-number {
      text-align: center;
      font-size: 14px;
      font-weight: 600;
      color: #333;
      margin-bottom: 10px;
    }
    .cards-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 10px;
      justify-items: center;
    }
    .borda-pontilhada {
      border: 1px dotted #ccc;
      padding: 5px;
      position: relative;
    }
    .carteirinha::before {
      content: "";
      position: absolute;
      top: -15px;
      left: 50%;
      transform: translateX(-50%);
      width: 60px;
      height: 30px;
      background-color: #d0d0d0;
      border-radius: 15px;
      z-index: 10;
    }
    .borda-pontilhada::after {
      content: "✂️";
      position: absolute;
      top: -14px;
      right: -11px;
      font-size: 16px;
      color: #2196F3;
    }
    input {
      width: 100%;
      padding: 12px 20px;
      margin: 8px 0;
      border: 1px solid #ccc;
      border-radius: 8px;
      box-sizing: border-box;
      font-size: 16px;
    }
    input:focus {
      border-color: #008CBA;
      box-shadow: 0 0 5px rgba(0, 140, 186, 0.5);
      outline: none;
    }
    .carteirinha {
      background-color: #fff;
      border-radius: 10px;
      box-shadow: 0 2px 8px rgba(0,0,0,0.1);
      overflow: hidden;
      display: flex;
      flex-direction: column;
      width: 270px;
      height: 550px;
      padding: 10px;
      position: relative;
      border: 4px solid #2196F3;
    }
    .escola {
      font-size: 16px;
      font-weight: 600;
      color: #2196F3;
      margin-bottom: 10px;
      text-align: center;
      text-transform: uppercase;
      letter-spacing: 1px;
      margin-top: 10px;
    }
    .foto {
      width: 140px;
      height: 160px;
      margin-bottom: 10px;
      border-radius: 50%;
      object-fit: cover;
      margin-left: auto;
      margin-right: auto;
      border: 4px solid #2196F3;
      cursor: pointer;
    }
    .info {
      display: flex;
      flex-direction: column;
      align-items: flex-start;
      text-align: left;
      margin-left: 15px;
      margin-bottom: 10px;
      font-size: 12px;
      color: #333;
    }
    .info div, .info span {
      margin: 3px 0;
    }
    .info .titulo {
      font-weight: 600;
      color: #2196F3;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    .info .descricao {
      color: #555;
    }
    .info .linha, .info .linha-ra {
      display: flex;
      justify-content: space-between;
      width: 100%;
    }
    .info .linha-ra {
      flex-wrap: nowrap;
      white-space: nowrap;
    }
    .linha-nome {
      display: flex;
      align-items: center;
      gap: 5px;
    }
    .status {
      padding: 6px;
      font-weight: 600;
      border-radius: 8px;
      color: #fff;
      text-transform: uppercase;
      margin-bottom: 5px;
      display: flex;
      justify-content: center;
      align-items: center;
      height: 35px;
      min-width: 150px;
      text-align: center;
    }
    .verde {
      background-color: #81C784;
    }
    .vermelho {
      background-color: #E57373;
    }
    .qr-code-container {
      display: flex;
      justify-content: center;
      align-items: center;
      flex-direction: column;
      margin-top: 10px;
    }
    .qr-code-text {
      font-size: 12px;
      font-weight: 600;
      margin-bottom: 2px;
      text-align: center;
    }
    .qr-code-container img {
      width: 80px;
      height: 80px;
    }
    /* Overlay de loading e mensagem de sucesso para a geração de carteirinhas */
    #loading-overlay {
      position: fixed;
      top: 0; left: 0; right: 0; bottom: 0;
      background: rgba(0, 0, 0, 0.5);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 9999;
    }
    #loading-overlay .spinner-border {
      width: 3rem;
      height: 3rem;
    }
    #cards-success {
      display: none;
      position: fixed;
      top: 10px;
      left: 50%;
      transform: translateX(-50%);
      background: #d4edda;
      color: #155724;
      padding: 10px;
      border-radius: 5px;
      z-index: 10000;
    }
    @media print {
      body {
        background-color: #fff;
        margin: 0;
        padding: 0;
      }
      #search-container {
        display: none !important;
      }
      .carteirinhas-container {
        width: 100%;
        max-width: 1100px;
      }
      .cards-grid {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 20px;
        justify-items: center;
      }
      .borda-pontilhada {
        border: none !important;
        padding: 10px !important;
        position: relative;
      }
      .borda-pontilhada::before {
        content: "";
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        pointer-events: none;
        background-image: radial-gradient(circle, #000 20%, transparent 20%);
        background-size: 10px 10px;
        border-radius: 4px;
      }
      .carteirinha::before {
        width: 60px !important;
        height: 30px !important;
        background-color: #d0d0d0 !important;
        border-radius: 15px !important;
        z-index: 10 !important;
        top: -15px !important;
        left: 50% !important;
        transform: translateX(-50%) !important;
      }
      .carteirinha {
        width: 250px;
        height: 455px !important;
        page-break-inside: avoid;
        padding-top: 5px;
      }
      .status {
        height: 30px;
        line-height: 30px;
        text-align: center;
      }
      .foto {
        width: 110px;
        height: 130px;
        border: 4px solid #2196F3;
        object-fit: cover;
      }
      .info {
        margin-top: 5px;
      }
      .page {
        page-break-after: always;
        break-after: page;
        margin-bottom: 0;
      }
      @page {
        size: A4 portrait;
        margin: 5mm;
      }
      .imprimir-carteirinhas, .imprimir-pagina {
        display: none !important;
      }
    }
    .imprimir-carteirinhas {
      position: fixed;
      bottom: 20px;
      right: 20px;
      background-color: #2196F3;
      color: #fff;
      padding: 10px 20px;
      font-size: 16px;
      border-radius: 8px;
      cursor: pointer;
      box-shadow: 0 4px 6px rgba(0,0,0,0.2);
    }
    .imprimir-pagina {
      background-color: #FF5722;
      color: #fff;
      padding: 10px 20px;
      font-size: 14px;
      border-radius: 5px;
      cursor: pointer;
      margin: 5px auto;
      display: block;
    }
    .imprimir-pagina:hover {
      background-color: #FF7043;
    }
    @media screen {
      .page {
        border: 2px dashed #ccc;
        padding: 10px;
        margin-bottom: 40px;
      }
    }
    .foto.uploadable:hover {
      transform: scale(1.05);
      transition: all 0.3s ease;
      cursor: pointer;
    }
  </style>
</head>
<body>
  <div id="loading-overlay">
    <div style="text-align: center; color: white;">
      <div class="spinner-border" role="status">
        <span class="sr-only">Carregando...</span>
      </div>
      <p>Carregando carteirinhas...</p>
    </div>
  </div>
  <div id="cards-success">Carteirinhas geradas com sucesso</div>
  <div class="carteirinhas-container">
    <button class="imprimir-carteirinhas" onclick="imprimirCarteirinhas()">Imprimir Carteirinhas</button>
    <div id="search-container">
      <input type="text" id="localizarAluno" placeholder="Localizar Aluno">
    </div>
"""
    contador = 0
    num_pagina = 1
    html_content += '<div class="page"><div class="page-number">Página ' + str(num_pagina) + '</div>'
    html_content += '<button class="imprimir-pagina" onclick="imprimirPagina(this)">Imprimir Página</button>'
    html_content += '<div class="cards-grid">'
    
    def gerar_qr_code(rm):
        caminho_qr = f'static/qr_codes/{rm}.png'
        qr = qrcode.make(f"RM: {rm} - Se possível, contribua com a APM")
        qr.save(caminho_qr)
        return f"/static/qr_codes/{rm}.png"
    
    for _, row in dados.iterrows():
        rm = str(row['RM'])
        if rm == '0':
            continue
        
        nome = row['NOME']
        data_nasc = row['DATA NASC.']
        serie = row['SÉRIE']
        horario = row['HORÁRIO']
        
        qr_url = gerar_qr_code(rm)
        
        if pd.notna(data_nasc):
            try:
                data_nasc = pd.to_datetime(data_nasc, errors='coerce')
                if pd.notna(data_nasc):
                    data_nasc = data_nasc.strftime('%d/%m/%Y')
                else:
                    data_nasc = "Desconhecida"
            except Exception as e:
                data_nasc = "Desconhecida"
        else:
            data_nasc = "Desconhecida"
        
        ra = row['RA']
        sai_sozinho = row['SAI SOZINHO?']
        if sai_sozinho == 'Sim':
            classe_cor = 'verde'
            status_texto = "Sai Sozinho"
        else:
            classe_cor = 'vermelho'
            status_texto = "Não Sai Sozinho"
        
        allowed_exts = ['.jpg', '.jpeg', '.png', '.bmp', '.gif']
        found_photo = None
        for ext in allowed_exts:
            caminho_foto = f'static/fotos/{rm}{ext}'
            if os.path.exists(caminho_foto):
                found_photo = f"/static/fotos/{rm}{ext}"
                break
        
        if found_photo:
            foto_tag = f'<img src="{found_photo}" alt="Foto" class="foto uploadable" data-rm="{rm}">'
        else:
            foto_tag = f'''
            <div class="foto uploadable" data-rm="{rm}" style="display:flex; flex-direction:column; align-items:center; justify-content:center;">
              <span style="font-size:40px; opacity:0.5; color: grey;">&#128247;</span>
              <small style="font-size:12px; opacity:0.5; color: grey;">Anexe uma foto</small>
            </div>
            '''
        
        hidden_input = f'<input type="file" class="inline-upload" data-rm="{rm}" style="display:none;" accept="image/*">'
        qr_tag = f'<img src="{qr_url}" alt="QR Code">'
        
        html_content += f"""
      <div class="borda-pontilhada">
        <div class="carteirinha">
          <div class="escola">E.M José Padin Mouta</div>
          {foto_tag}
          {hidden_input}
          <div class="info">
            <div class="linha-nome">
              <span class="titulo">Nome:</span>
              <span class="descricao">{nome}</span>
            </div>
            <div class="linha-rm" style="white-space: nowrap;">
              <span class="titulo">RM:</span>
              <span class="descricao">{rm}</span>
            </div>
            <div class="linha">
              <div class="titulo">Série:</div>
              <div class="descricao">{serie}</div>
              <div class="titulo">Data Nasc.:</div>
              <div class="descricao">{data_nasc}</div>
            </div>
            <div class="linha-ra">
              <div class="titulo">RA:</div>
              <div class="descricao">{ra}</div>
              <div class="titulo">Horário:</div>
              <div class="descricao">{horario}</div>
            </div>
          </div>
          <div class="status {classe_cor}">{status_texto}</div>
          <div class="qr-code-container">
            <div class="qr-code-text">Se possível, contribua com a APM</div>
            {qr_tag}
          </div>
        </div>
      </div>
"""
        contador += 1
        if contador % 4 == 0:
            html_content += '</div></div>'
            if contador < len(dados):
                num_pagina += 1
                html_content += '<div class="page"><div class="page-number">Página ' + str(num_pagina) + '</div>'
                html_content += '<button class="imprimir-pagina" onclick="imprimirPagina(this)">Imprimir Página</button>'
                html_content += '<div class="cards-grid">'
    
    if contador % 4 != 0:
        html_content += '</div></div>'
    
    html_content += """
  </div>
  <script>
    // Remoção do overlay de loading e exibição da mensagem de sucesso após o carregamento completo
    window.onload = function(){
      var overlay = document.getElementById('loading-overlay');
      overlay.style.display = 'none';
      var cardsMsg = document.getElementById('cards-success');
      cardsMsg.style.display = 'block';
      setTimeout(function(){
        cardsMsg.style.display = 'none';
      }, 3000);
    };

    function imprimirCarteirinhas() {
      window.print();
    }
    function imprimirPagina(botao) {
      let pagina = botao.closest('.page');
      let todasPaginas = document.querySelectorAll('.page');
      todasPaginas.forEach(p => {
        if (p !== pagina) {
          p.style.display = 'none';
        }
      });
      setTimeout(() => {
        window.print();
        todasPaginas.forEach(p => {
          p.style.display = '';
        });
      }, 100);
    }
    document.getElementById('localizarAluno').addEventListener('keyup', function() {
      var filtro = this.value.toLowerCase();
      var cards = document.querySelectorAll('.borda-pontilhada');
      cards.forEach(function(card) {
        var nomeElem = card.querySelector('.linha-nome .descricao');
        if (nomeElem) {
          var nome = nomeElem.textContent.toLowerCase();
          if (nome.indexOf(filtro) > -1) {
            card.style.display = '';
          } else {
            card.style.display = 'none';
          }
        }
      });
    });
    
    // Variável para controlar o tempo da mensagem de upload inline
    var flashTimeout = null;
    document.addEventListener('DOMContentLoaded', function() {
      document.querySelectorAll('.uploadable').forEach(function(element) {
        element.addEventListener('click', function() {
          var rm = element.getAttribute('data-rm');
          var input = document.querySelector('.inline-upload[data-rm="'+rm+'"]');
          if(input) {
            input.click();
          }
        });
      });
      
      document.querySelectorAll('.inline-upload').forEach(function(input) {
        input.addEventListener('change', function() {
          var file = input.files[0];
          if(file) {
            var rm = input.getAttribute('data-rm');
            var formData = new FormData();
            formData.append('rm', rm);
            formData.append('foto_file', file);
            
            fetch('/upload_inline_foto', {
              method: 'POST',
              body: formData
            })
            .then(response => response.json())
            .then(data => {
              if(data.url) {
                var uploadable = document.querySelector('.uploadable[data-rm="'+rm+'"]');
                if(uploadable.tagName.toLowerCase() === 'img') {
                  uploadable.src = data.url;
                } else {
                  var img = document.createElement('img');
                  img.src = data.url;
                  img.alt = "Foto";
                  img.className = "foto uploadable";
                  img.setAttribute('data-rm', rm);
                  uploadable.parentNode.replaceChild(img, uploadable);
                }
                // Exibe mensagem de sucesso para upload inline
                var msgDiv = document.getElementById('upload-success');
                if(!msgDiv){
                  msgDiv = document.createElement('div');
                  msgDiv.id = 'upload-success';
                  msgDiv.style.position = 'fixed';
                  msgDiv.style.top = '10px';
                  msgDiv.style.right = '10px';
                  msgDiv.style.backgroundColor = '#d4edda';
                  msgDiv.style.color = '#155724';
                  msgDiv.style.padding = '10px';
                  msgDiv.style.borderRadius = '5px';
                  document.body.appendChild(msgDiv);
                }
                msgDiv.style.display = 'block';
                msgDiv.innerHTML = data.message;
                if(flashTimeout) {
                  clearTimeout(flashTimeout);
                }
                flashTimeout = setTimeout(function(){
                  msgDiv.style.display = 'none';
                }, 3000);
              } else {
                alert("Erro ao fazer upload: " + (data.error || "Erro desconhecido"));
              }
            })
            .catch(error => {
              console.error('Erro:', error);
              alert("Erro no upload da foto.");
            });
          }
        });
      });
    });
  </script>
</body>
</html>
"""
    return html_content

@app.route('/login', methods=['GET', 'POST'])
def login():
    error = None
    if request.method == 'POST':
        token = request.form.get('token')
        if token == ACCESS_TOKEN:
            session['logged_in'] = True
            return redirect(url_for('index'))
        else:
            error = "Token inválido. Tente novamente."
    login_html = '''
    <!doctype html>
    <html lang="pt-br">
      <head>
        <meta charset="utf-8">
        <title>Login - Acesso Restrito</title>
        <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
        <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
        <style>
          body {
            background: linear-gradient(135deg, #283E51, #4B79A1);
            font-family: 'Montserrat', sans-serif;
            margin: 0;
            padding: 0;
            display: flex;
            flex-direction: column;
            min-height: 100vh;
          }
          header {
            background-color: #283E51;
            color: #fff;
            padding: 20px;
            text-align: center;
          }
          main {
            flex: 1;
            display: flex;
            align-items: center;
            justify-content: center;
          }
          .container-login {
            background: #fff;
            padding: 40px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            width: 100%;
            max-width: 400px;
          }
          .container-login h2 {
            margin-bottom: 20px;
            font-weight: 600;
            color: #283E51;
          }
          .btn-primary {
            background-color: #283E51;
            border: none;
          }
          .btn-primary:hover {
            background-color: #1d2d3a;
          }
          footer {
            background-color: #424242;
            color: #fff;
            text-align: center;
            padding: 10px;
          }
          .error {
            color: red;
            margin-top: 15px;
          }
        </style>
      </head>
      <body>
        <header>
          <h1 class="mb-0">Carteirinhas - E.M José Padin Mouta</h1>
        </header>
        <main>
          <div class="container container-login">
            <h2>Acesso Restrito</h2>
            <form method="POST">
              <div class="form-group">
                <input type="password" name="token" class="form-control" placeholder="Digite o token de acesso" required>
              </div>
              <button type="submit" class="btn btn-primary btn-block">Entrar</button>
            </form>
            {% if error %}
              <p class="error">{{ error }}</p>
            {% endif %}
          </div>
        </main>
        <footer>
          Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
        </footer>
      </body>
    </html>
    '''
    return render_template_string(login_html, error=error)

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route('/', methods=['GET', 'POST'])
@login_required
def index():
    if request.method == 'POST':
        if 'excel_file' in request.files:
            file = request.files['excel_file']
            if file.filename == '':
                return "Nenhum arquivo selecionado", 400
            # Exibe uma mensagem de loading ao clicar em "Gerar Carteirinhas"
            flash("Gerando carteirinhas. Aguarde...", "info")
            html_result = gerar_html_carteirinhas(file)
            # Após gerar, adicione uma mensagem de sucesso via JS no template gerado (veja o overlay no código acima)
            return html_result
    index_html = '''
    <!doctype html>
    <html lang="pt-br">
      <head>
        <meta charset="utf-8">
        <title>Upload para Carteirinhas</title>
        <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
        <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
        <style>
          body {
            background: #eef2f3;
            font-family: 'Montserrat', sans-serif;
          }
          header {
            background-color: #283E51;
            color: #fff;
            padding: 20px;
            text-align: center;
          }
          .container-upload {
            background: #fff;
            padding: 40px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            margin: 40px auto;
            max-width: 800px;
          }
          h2 {
            color: #283E51;
            font-weight: 600;
          }
          .btn-primary {
            background-color: #283E51;
            border: none;
          }
          .btn-primary:hover {
            background-color: #1d2d3a;
          }
          .btn-secondary {
            background-color: #4B79A1;
            border: none;
          }
          .btn-secondary:hover {
            background-color: #3a5d78;
          }
          .logout-container {
            text-align: center;
            margin-top: 20px;
          }
          .btn-logout {
            background-color: #dc3545;
            color: #fff;
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            font-size: 16px;
            text-decoration: none;
            transition: background-color 0.3s;
          }
          .btn-logout:hover {
            background-color: #c82333;
          }
          footer {
            background-color: #424242;
            color: #fff;
            text-align: center;
            padding: 10px;
            position: fixed;
            bottom: 0;
            width: 100%;
          }
          #multi-upload-section {
            margin-top: 20px;
            border: 1px solid #ccc;
            padding: 20px;
            border-radius: 8px;
            background-color: #f9f9f9;
          }
          .multi-upload-group {
            margin-bottom: 15px;
          }
          /* Oculta as mensagens flash após 3 segundos */
          #flash-messages {
            position: fixed;
            top: 10px;
            left: 50%;
            transform: translateX(-50%);
            z-index: 10000;
          }
        </style>
      </head>
      <body>
        <header>
          <h1 class="mb-0">Carteirinhas - E.M José Padin Mouta</h1>
        </header>
        <div class="container container-upload">
          {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
              <div id="flash-messages">
                {% for category, message in messages %}
                  <div class="alert alert-{{ 'success' if category == 'success' else 'info' }}" role="alert">{{ message }}</div>
                {% endfor %}
              </div>
            {% endif %}
          {% endwith %}
          <h2 class="mb-4">Envie a lista piloto (Excel)</h2>
          <form method="POST" enctype="multipart/form-data" onsubmit="showLoading()">
            <div class="form-group">
              <input type="file" class="form-control-file" name="excel_file" accept=".xlsx, .xls">
            </div>
            <button type="submit" class="btn btn-primary">Gerar Carteirinhas</button>
          </form>
          <hr>
          <h2 class="mb-4">Upload da Foto</h2>
          <form method="POST" action="/upload_foto" enctype="multipart/form-data">
            <div class="form-group">
              <label>RM do Aluno:</label>
              <input type="text" class="form-control" name="rm" placeholder="Digite o RM">
            </div>
            <div class="form-group">
              <input type="file" class="form-control-file" name="foto_file" accept="image/*">
            </div>
            <button type="submit" class="btn btn-secondary">Enviar Foto</button>
          </form>
          <hr>
          <h2 class="mb-4">Upload de Múltiplas Fotos</h2>
          <button type="button" class="btn btn-secondary" id="show-multi-upload">Enviar múltiplas fotos</button>
          <div id="multi-upload-section" style="display: none;">
            <form method="POST" action="/upload_multiplas_fotos" enctype="multipart/form-data" id="multi-upload-form">
              <div id="multi-upload-fields">
                <div class="multi-upload-group">
                  <div class="form-group">
                    <label>RM do Aluno:</label>
                    <input type="text" class="form-control" name="rm[]" placeholder="Digite o RM">
                  </div>
                  <div class="form-group">
                    <input type="file" class="form-control-file" name="foto_file[]" accept="image/*">
                  </div>
                </div>
              </div>
              <button type="button" class="btn btn-info" id="add-more" style="margin-top:10px;">Adicionar outra foto</button>
              <button type="submit" class="btn btn-primary" style="margin-top:10px;">Enviar Fotos</button>
            </form>
          </div>
          <div class="logout-container">
            <a href="/logout" class="btn-logout">
              <i class="fas fa-sign-out-alt"></i> Logout
            </a>
          </div>
        </div>
        <footer>
          Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
        </footer>
        <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@4.5.2/dist/js/bootstrap.bundle.min.js"></script>
        <script>
          // Remove as mensagens flash após 3 segundos
          setTimeout(function(){
            var flashDiv = document.getElementById('flash-messages');
            if(flashDiv){
              flashDiv.style.display = 'none';
            }
          }, 3000);
          
          function showLoading(){
            // Exibe uma sobreposição de loading para geração de carteirinhas
            var loadingOverlay = document.createElement('div');
            loadingOverlay.id = 'loading-overlay';
            loadingOverlay.style.position = 'fixed';
            loadingOverlay.style.top = '0';
            loadingOverlay.style.left = '0';
            loadingOverlay.style.right = '0';
            loadingOverlay.style.bottom = '0';
            loadingOverlay.style.background = 'rgba(0,0,0,0.5)';
            loadingOverlay.style.display = 'flex';
            loadingOverlay.style.alignItems = 'center';
            loadingOverlay.style.justifyContent = 'center';
            loadingOverlay.style.zIndex = '9999';
            loadingOverlay.innerHTML = `
              <div style="text-align: center; color: white;">
                <div class="spinner-border" role="status">
                  <span class="sr-only">Carregando...</span>
                </div>
                <p>Carregando carteirinhas...</p>
              </div>
            `;
            document.body.appendChild(loadingOverlay);
          }
          
          document.getElementById('show-multi-upload').addEventListener('click', function() {
            var section = document.getElementById('multi-upload-section');
            if(section.style.display === 'none') {
              section.style.display = 'block';
            } else {
              section.style.display = 'none';
            }
          });
          document.getElementById('add-more').addEventListener('click', function() {
            var container = document.getElementById('multi-upload-fields');
            var group = document.createElement('div');
            group.className = 'multi-upload-group';
            group.innerHTML = `
              <div class="form-group">
                <label>RM do Aluno:</label>
                <input type="text" class="form-control" name="rm[]" placeholder="Digite o RM">
              </div>
              <div class="form-group">
                <input type="file" class="form-control-file" name="foto_file[]" accept="image/*">
              </div>
            `;
            container.appendChild(group);
          });
        </script>
      </body>
    </html>
    '''
    return render_template_string(index_html)

@app.route('/upload_foto', methods=['POST'])
def upload_foto():
    if 'foto_file' not in request.files:
        return "Nenhum arquivo de foto enviado", 400
    rm = request.form.get('rm')
    if not rm:
        return "RM não fornecido", 400
    file = request.files['foto_file']
    if file.filename == '':
        return "Nenhuma foto selecionada", 400
    if not allowed_file(file.filename):
        return "Formato de imagem não permitido", 400
    
    original_filename = secure_filename(file.filename)
    _, ext = os.path.splitext(original_filename)
    new_filename = secure_filename(f"{rm}{ext.lower()}")
    file_path = os.path.join('static', 'fotos', new_filename)
    file.save(file_path)
    flash("Foto anexada com sucesso", "success")
    return redirect(url_for('index'))

@app.route('/upload_multiplas_fotos', methods=['POST'])
def upload_multiplas_fotos():
    rms = request.form.getlist("rm[]")
    files = request.files.getlist("foto_file[]")
    
    if not files:
       return "Nenhuma foto enviada", 400
    
    for rm, file in zip(rms, files):
        if file.filename == '':
            continue
        if not rm:
            continue
        if not allowed_file(file.filename):
            continue
        original_filename = secure_filename(file.filename)
        _, ext = os.path.splitext(original_filename)
        new_filename = secure_filename(f"{rm}{ext.lower()}")
        file_path = os.path.join('static', 'fotos', new_filename)
        file.save(file_path)
    flash("Foto(s) anexada(s) com sucesso", "success")
    return redirect(url_for('index'))

@app.route('/upload_inline_foto', methods=['POST'])
def upload_inline_foto():
    if 'foto_file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400
    rm = request.form.get('rm')
    if not rm:
        return jsonify({'error': 'RM não fornecido'}), 400
    file = request.files['foto_file']
    if file.filename == '':
        return jsonify({'error': 'Nenhuma foto selecionada'}), 400
    if not allowed_file(file.filename):
        return jsonify({'error': 'Formato de imagem não permitido'}), 400
    original_filename = secure_filename(file.filename)
    _, ext = os.path.splitext(original_filename)
    new_filename = secure_filename(f"{rm}{ext.lower()}")
    file_path = os.path.join('static', 'fotos', new_filename)
    file.save(file_path)
    return jsonify({'url': f"/static/fotos/{new_filename}", 'message': "Foto anexada com sucesso"})

if __name__ == '__main__':
    app.run(debug=True)
