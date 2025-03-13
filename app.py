from flask import Flask, request, redirect, url_for, render_template_string, jsonify, session, flash, send_file       
import pandas as pd
import os
import uuid
import re
from datetime import datetime
from io import BytesIO
from werkzeug.utils import secure_filename
from functools import wraps
import locale
import xlrd  # Para ler arquivos .xls
from openpyxl import load_workbook, Workbook  # Usado para trabalhar com XLSX e para converter XLS

# Tenta definir a localidade para formatação de datas em português
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    pass

app = Flask(__name__)
app.secret_key = 'sua_chave_secreta'  # Altere para uma chave segura
ACCESS_TOKEN = "minha_senha"  # Token de acesso

app.config['UPLOAD_FOLDER'] = 'uploads'
ALLOWED_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif'}

# Cria os diretórios necessários, se não existirem
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

# Função para atualizar valor em célula mesclada
def set_merged_cell_value(ws, cell_coord, value):
    for merged_range in ws.merged_cells.ranges:
        if cell_coord in merged_range:
            range_str = str(merged_range)
            ws.unmerge_cells(range_str)
            ws[cell_coord] = value
            ws.merge_cells(range_str)
            return
    ws[cell_coord] = value

def convert_xls_to_xlsx(file_like):
    """
    Converte um arquivo XLS (file-like) para um Workbook do openpyxl.
    """
    # Lê o conteúdo do XLS usando xlrd
    book_xlrd = xlrd.open_workbook(file_contents=file_like.read())
    wb = Workbook()
    # Remove a planilha padrão se houver e houver planilhas no arquivo XLS
    if "Sheet" in wb.sheetnames and len(book_xlrd.sheet_names()) > 0:
        std = wb.active
        wb.remove(std)
    for sheet_name in book_xlrd.sheet_names():
        sheet_xlrd = book_xlrd.sheet_by_name(sheet_name)
        ws = wb.create_sheet(title=sheet_name)
        for row in range(sheet_xlrd.nrows):
            for col in range(sheet_xlrd.ncols):
                ws.cell(row=row+1, column=col+1, value=sheet_xlrd.cell_value(row, col))
    return wb

def load_workbook_model(file):
    """
    Abre o arquivo do modelo. Se for XLSX, utiliza openpyxl.load_workbook;
    se for XLS, converte para um Workbook do openpyxl.
    """
    ext = os.path.splitext(file.filename)[1].lower()
    file.seek(0)
    if ext == '.xlsx':
        return load_workbook(file)
    elif ext == '.xls':
        content = file.read()
        return convert_xls_to_xlsx(BytesIO(content))
    else:
        raise ValueError("Formato de arquivo não suportado para o quadro modelo.")

def gerar_html_carteirinhas(arquivo_excel):
    planilha = pd.read_excel(arquivo_excel, sheet_name='LISTA CORRIDA')
    dados = planilha[['RM', 'NOME', 'DATA NASC.', 'RA', 'SAI SOZINHO?', 'SÉRIE', 'HORÁRIO']]
    dados['RM'] = dados['RM'].fillna(0).astype(int)
    
    alunos_sem_fotos_list = []
    html_content = """
<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Carteirinhas - E.M José Padin Mouta</title>
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
    #search-container { margin-top: 10px; }
    #localizarAluno { padding: 0.2cm; font-size: 0.3cm; width: 3.5cm; }
    .carteirinhas-container { width: 100%; max-width: 1100px; }
    .page { margin-bottom: 40px; position: relative; }
    .page-number { text-align: center; font-size: 0.3cm; font-weight: 600; color: #333; margin-bottom: 0.2cm; }
    .cards-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 0.2cm; justify-items: center; }
    .borda-pontilhada { border: 0.05cm dotted #ccc; padding: 0.1cm; position: relative; }
    .borda-pontilhada::after {
      content: "✂️";
      position: absolute;
      top: -0.35cm;
      right: -0.30cm;
      font-size: 0.3cm;
      color: #2196F3;
    }
    input {
      width: 100%;
      padding: 0.2cm;
      margin: 0.1cm 0;
      border: 0.05cm solid #ccc;
      border-radius: 0.2cm;
      box-sizing: border-box;
      font-size: 0.3cm;
    }
    input:focus {
      border-color: #008CBA;
      box-shadow: 0 0 0.2cm rgba(0, 140, 186, 0.5);
      outline: none;
    }
    .carteirinha {
      background-color: #fff;
      border-radius: 0.3cm;
      box-shadow: 0 0.1cm 0.2cm rgba(0,0,0,0.1);
      overflow: hidden;
      display: flex;
      flex-direction: column;
      width: 6.0cm;
      height: 9.0cm;
      padding: 0.2cm;
      position: relative;
      border: 0.05cm solid #2196F3;
    }
    .escola {
      font-size: 0.35cm;
      font-weight: 600;
      color: #2196F3;
      margin-bottom: 0.1cm;
      text-align: center;
      text-transform: uppercase;
      letter-spacing: 0.05cm;
      margin-top: 0.1cm;
      white-space: nowrap;
    }
    .foto {
      width: 1.8cm;
      height: 1.8cm;
      margin-bottom: 0.1cm;
      border-radius: 50%;
      object-fit: cover;
      margin-left: auto;
      margin-right: auto;
      border: 0.1cm solid #2196F3;
      cursor: pointer;
    }
    .info {
      display: flex;
      flex-direction: column;
      align-items: flex-start;
      text-align: left;
      margin-left: 0.1cm;
      margin-bottom: 0.1cm;
      font-size: 0.3cm;
      color: #333;
    }
    .info div, .info span { margin: 0.08cm 0; }
    .info .titulo {
      font-weight: 600;
      color: #2196F3;
      text-transform: uppercase;
      letter-spacing: 0.02cm;
    }
    .info .descricao { color: #555; }
    .linha-nome { display: flex; align-items: center; gap: 0.1cm; }
    .linha, .linha-ra, .linha-horario, .linha-rm { display: flex; flex-direction: row; align-items: center; gap: 0.2cm; }
    .status {
      padding: 0.2cm;
      font-weight: 600;
      border-radius: 0.2cm;
      color: #fff;
      text-transform: uppercase;
      margin-bottom: 0.1cm;
      display: flex;
      justify-content: center;
      align-items: center;
      height: 0.6cm;
      min-width: 1.5cm;
      text-align: center;
    }
    .verde { background-color: #81C784; }
    .vermelho { background-color: #E57373; }
    .ano { position: absolute; bottom: 0.2cm; left: 0; right: 0; text-align: center; font-size: 0.4cm; font-weight: 600; color: #2196F3; }
    #loading-overlay {
      position: fixed;
      top: 0; left: 0; right: 0; bottom: 0;
      background: rgba(0, 0, 0, 0.5);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 9999;
    }
    #cards-success {
      display: none;
      position: fixed;
      top: 10px;
      left: 50%;
      transform: translateX(-50%);
      background: #d4edda;
      color: #155724;
      padding: 0.2cm;
      border-radius: 0.2cm;
      z-index: 10000;
    }
    .no-print { }
    @media print {
      .no-print { display: none !important; }
      body {
        margin: 0;
        padding: 0;
        font-size: 16px;
      }
    }
    .imprimir-carteirinhas {
      position: fixed;
      bottom: 0.5cm;
      right: 0.5cm;
      background-color: #2196F3;
      color: #fff;
      padding: 0.2cm 0.4cm;
      font-size: 0.3cm;
      border-radius: 0.2cm;
      cursor: pointer;
      box-shadow: 0 0.1cm 0.2cm rgba(0,0,0,0.2);
    }
    .imprimir-pagina {
      background-color: #FF5722;
      color: #fff;
      padding: 0.2cm 0.4cm;
      font-size: 0.3cm;
      border-radius: 0.2cm;
      cursor: pointer;
      margin: 0.2cm auto;
      display: block;
    }
    .imprimir-pagina:hover { background-color: #FF7043; }
    .alunos-sem-fotos-btn {
      background-color: #4B79A1;
      color: #fff;
      border: none;
      padding: 0.2cm 0.5cm;
      font-size: 0.3cm;
      border-radius: 0.2cm;
      cursor: pointer;
      margin-bottom: 0.2cm;
    }
    .alunos-sem-fotos-btn:hover { background-color: #3a5d78; }
    #relatorio-container {
      display: none;
      position: fixed;
      top: 10%;
      left: 50%;
      transform: translateX(-50%);
      width: 80%;
      max-height: 80%;
      overflow-y: auto;
      background: #fff;
      border: 1px solid #ccc;
      border-radius: 10px;
      box-shadow: 0 0 10px rgba(0,0,0,0.5);
      z-index: 10000;
      padding: 20px;
    }
    #relatorio-container h2 { text-align: center; margin-top: 0; }
    #relatorio-container table {
      width: 100%;
      border-collapse: collapse;
    }
    #relatorio-container th, #relatorio-container td {
      border: 1px solid #ccc;
      padding: 8px;
      text-align: left;
    }
    #relatorio-container button.close-relatorio {
      float: right;
      font-size: 1.2em;
      border: none;
      background: none;
      cursor: pointer;
    }
    /* Novo estilo para o header, mais suave e moderno */
    header {
      background: linear-gradient(90deg, #283E51, #4B79A1);
      color: #fff;
      padding: 20px;
      text-align: center;
      border-bottom: 3px solid #1d2d3a;
      border-radius: 0 0 15px 15px;
      box-shadow: 0 4px 6px rgba(0,0,0,0.1);
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
    <div class="no-print" style="margin-bottom: 10px;">
      <button class="alunos-sem-fotos-btn" onclick="mostrarRelatorioAlunosSemFotos()">Alunos sem fotos</button>
      <button class="imprimir-carteirinhas" onclick="imprimirCarteirinhas()">Imprimir Carteirinhas</button>
    </div>
    <div id="search-container">
      <input type="text" id="localizarAluno" placeholder="Localizar Aluno">
    </div>
"""
    contador = 0
    num_pagina = 1
    html_content += '<div class="page"><div class="page-number">Página ' + str(num_pagina) + '</div>'
    html_content += '<button class="imprimir-pagina" onclick="imprimirPagina(this)">Imprimir Página</button>'
    html_content += '<div class="cards-grid">'
    
    for _, row in dados.iterrows():
        if not str(row['RM']).strip() or str(row['RM']).strip() == "0":
            continue
        
        nome = row['NOME']
        data_nasc = row['DATA NASC.']
        serie = row['SÉRIE']
        horario = row['HORÁRIO']
        
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
            caminho_foto = f'static/fotos/{row["RM"]}{ext}'
            if os.path.exists(caminho_foto):
                found_photo = f"/static/fotos/{row['RM']}{ext}"
                break
        
        if not found_photo:
            alunos_sem_fotos_list.append({
                'rm': row['RM'],
                'nome': nome,
                'serie': serie
            })
        
        if found_photo:
            foto_tag = f'<img src="{found_photo}" alt="Foto" class="foto uploadable" data-rm="{row["RM"]}">'
        else:
            foto_tag = f'''
            <div class="foto uploadable" data-rm="{row["RM"]}" style="display: flex; flex-direction: column; align-items: center; justify-content: center;">
              <span style="font-size:0.8cm; opacity:0.5; color: grey; margin-bottom: 0.1cm;">&#128247;</span>
              <small style="font-size:0.2cm; opacity:0.5; color: grey;">Anexe uma foto</small>
            </div>
            '''
        
        hidden_input = f'<input type="file" class="inline-upload" data-rm="{row["RM"]}" style="display:none;" accept="image/*">'
        
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
            <div class="linha-rm">
              <span class="titulo">RM:</span>
              <span class="descricao">{row['RM']}</span>
            </div>
            <div class="linha">
              <div class="titulo">Série:</div>
              <div class="descricao">{serie}</div>
            </div>
            <div class="linha">
              <div class="titulo">Data Nasc.:</div>
              <div class="descricao">{data_nasc}</div>
            </div>
            <div class="linha-ra">
              <span class="titulo">RA:</span>
              <span class="descricao">{ra}</span>
            </div>
            <div class="linha-horario">
              <div class="titulo">Horário:</div>
              <div class="descricao">{horario}</div>
            </div>
          </div>
          <div class="status {classe_cor}">{status_texto}</div>
          <div class="ano">2025</div>
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
    
    relatorio_linhas = ""
    for aluno in alunos_sem_fotos_list:
        relatorio_linhas += f"<tr><td>{aluno['rm']}</td><td>{aluno['nome']}</td><td>{aluno['serie']}</td></tr>"
    
    html_content += f"""
  </div>
  <div id="relatorio-container" class="no-print">
    <button class="close-relatorio" onclick="fecharRelatorio()">&times;</button>
    <h2>Alunos sem Fotos</h2>
    <table>
      <thead>
        <tr>
          <th>RM</th>
          <th>Nome</th>
          <th>Série</th>
        </tr>
      </thead>
      <tbody>
        {relatorio_linhas}
      </tbody>
    </table>
  </div>
  <script>
  function showLoading() {{
    var existingOverlay = document.getElementById('loading-overlay');
    if (existingOverlay) {{
      existingOverlay.remove();
    }}
  
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
      <div style="text-align: center; color: white; font-family: Arial, sans-serif;">
        <svg width="3.0cm" height="4.5cm" viewBox="0 0 6.0 9.0" xmlns="http://www.w3.org/2000/svg">
          <rect x="0.3" y="0.3" width="5.4" height="8.4" rx="0.3" ry="0.3" stroke="white" stroke-width="0.1" fill="none" />
          <rect id="badge-fill" x="0.3" y="8.7" width="5.4" height="0" rx="0.3" ry="0.3" fill="white" />
        </svg>
        <p id="loading-text" style="margin-top: 0.2cm;">Gerando carteirinhas...</p>
      </div>
    `;
  
    document.body.appendChild(loadingOverlay);
  
    let fillHeight = 0;
    const maxHeight = 8.4; 
    function animateBadge() {{
      fillHeight += 0.2;
      if (fillHeight > maxHeight) {{
        fillHeight = maxHeight;
        clearInterval(interval);
      }}
      const badgeFill = document.getElementById('badge-fill');
      badgeFill.setAttribute('y', 8.7 - fillHeight);
      badgeFill.setAttribute('height', fillHeight);
    }}
  
    var interval = setInterval(animateBadge, 100);
    loadingOverlay.dataset.animationId = interval;
  }}
  
  showLoading();
  
  window.onload = function() {{
    var overlay = document.getElementById('loading-overlay');
    if (overlay) {{
      var animationId = Number(overlay.dataset.animationId);
      clearInterval(animationId);
      overlay.style.display = 'none';
    }}
    var cardsMsg = document.getElementById('cards-success');
    if (cardsMsg) {{
      cardsMsg.style.display = 'block';
      cardsMsg.innerHTML = 'Carteirinhas geradas com sucesso!';
      setTimeout(function() {{
        cardsMsg.style.display = 'none';
      }}, 3000);
    }}
  }};
  
  function imprimirCarteirinhas() {{
    window.print();
  }}
  
  function imprimirPagina(botao) {{
    let pagina = botao.closest('.page');
    let todasPaginas = document.querySelectorAll('.page');
    todasPaginas.forEach(p => {{
      if (p !== pagina) {{
        p.style.display = 'none';
      }}
    }});
    setTimeout(() => {{
      window.print();
      todasPaginas.forEach(p => {{
        p.style.display = '';
      }});
    }}, 100);
  }}
  
  function mostrarRelatorioAlunosSemFotos() {{
    document.getElementById('relatorio-container').style.display = 'block';
  }}
  
  function fecharRelatorio() {{
    document.getElementById('relatorio-container').style.display = 'none';
  }}
  
  document.getElementById('localizarAluno').addEventListener('keyup', function() {{
    var filtro = this.value.toLowerCase();
    var cards = document.querySelectorAll('.borda-pontilhada');
    cards.forEach(function(card) {{
      var nomeElem = card.querySelector('.linha-nome .descricao');
      if (nomeElem) {{
        var nome = nomeElem.textContent.toLowerCase();
        if (nome.indexOf(filtro) > -1) {{
          card.style.display = '';
        }} else {{
          card.style.display = 'none';
        }}
      }}
    }});
  }});
  
  var flashTimeout = null;
  document.addEventListener('DOMContentLoaded', function() {{
    document.querySelectorAll('.uploadable').forEach(function(element) {{
      element.addEventListener('click', function() {{
        var rm = element.getAttribute('data-rm');
        var input = document.querySelector('.inline-upload[data-rm="'+rm+'"]');
        if(input) {{
          input.click();
        }}
      }});
    }});
    
    document.querySelectorAll('.inline-upload').forEach(function(input) {{
      input.addEventListener('change', function() {{
        var file = input.files[0];
        if(file) {{
          var rm = input.getAttribute('data-rm');
          var formData = new FormData();
          formData.append('rm', rm);
          formData.append('foto_file', file);
          
          fetch('/upload_inline_foto', {{
            method: 'POST',
            body: formData
          }})
          .then(response => response.json())
          .then(data => {{
            if(data.url) {{
              var uploadable = document.querySelector('.uploadable[data-rm="'+rm+'"]');
              if(uploadable.tagName.toLowerCase() === 'img') {{
                uploadable.src = data.url;
              }} else {{
                var img = document.createElement('img');
                img.src = data.url;
                img.alt = "Foto";
                img.className = "foto uploadable";
                img.setAttribute('data-rm', rm);
                uploadable.parentNode.replaceChild(img, uploadable);
              }}
              var msgDiv = document.getElementById('upload-success');
              if(!msgDiv) {{
                msgDiv = document.createElement('div');
                msgDiv.id = 'upload-success';
                msgDiv.style.position = 'fixed';
                msgDiv.style.top = '0.2cm';
                msgDiv.style.right = '0.2cm';
                msgDiv.style.backgroundColor = '#d4edda';
                msgDiv.style.color = '#155724';
                msgDiv.style.padding = '0.2cm';
                msgDiv.style.borderRadius = '0.2cm';
                document.body.appendChild(msgDiv);
              }}
              msgDiv.style.display = 'block';
              msgDiv.innerHTML = data.message;
              if(flashTimeout) {{
                clearTimeout(flashTimeout);
              }}
              flashTimeout = setTimeout(function() {{
                msgDiv.style.display = 'none';
              }}, 3000);
            }} else {{
              alert("Erro ao fazer upload: " + (data.error || "Erro desconhecido"));
            }}
          }})
          .catch(error => {{
            console.error('Erro:', error);
            alert("Erro no upload da foto.");
          }});
        }}
      }});
    }});
  }});
  </script>
</body>
</html>
"""
    return render_template_string(html_content)

def gerar_declaracao_escolar(file_path, rm, tipo):
    planilha = pd.read_excel(file_path, sheet_name='LISTA CORRIDA')
    def format_rm(x):
        try:
            return str(int(float(x)))
        except:
            return str(x)
    planilha['RM_str'] = planilha['RM'].apply(format_rm)
    try:
        rm_num = str(int(float(rm)))
    except:
        rm_num = str(rm)
    aluno = planilha[planilha['RM_str'] == rm_num]
    if aluno.empty:
        return None
    row = aluno.iloc[0]
    nome = row['NOME']
    serie = row['SÉRIE']
    if isinstance(serie, str):
        serie = re.sub(r"(\d+º)([A-Za-z])", r"\1 ano \2", serie)
    data_nasc = row['DATA NASC.']
    ra = row['RA']
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
    
    now = datetime.now()
    meses = {1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
             7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"}
    mes = meses[now.month].capitalize()
    data_extenso = f"Praia Grande, {now.day:02d} de {mes} de {now.year}"
    
    additional_css = '''
.print-button {
  background-color: #283E51;
  color: #fff;
  border: none;
  padding: 10px 20px;
  border-radius: 5px;
  cursor: pointer;
  margin-top: 20px;
}
.print-button:hover {
  background-color: #1d2d3a;
}
'''
    if tipo == "Escolaridade":
        titulo = "Declaração de Escolaridade"
        declaracao_text = (
            f"Declaro, para os devidos fins, que o(a) aluno(a) <strong><u>{nome}</u></strong>, "
            f"portador(a) do RA <strong><u>{ra}</u></strong>, nascido(a) em <strong><u>{data_nasc}</u></strong>, "
            f"encontra-se regularmente matriculado(a) na E.M José Padin Mouta, cursando atualmente o "
            f"<strong><u>{serie}</u></strong>."
        )
    elif tipo == "Transferencia":
        titulo = "Declaração de Transferência"
        declaracao_text = (
            f"Declaro, para os devidos fins, que o(a) aluno(a) <strong><u>{nome}</u></strong>, "
            f"portador(a) do RA <strong><u>{ra}</u></strong>, nascido(a) em <strong><u>{data_nasc}</u></strong>, "
            f"solicitou transferência de nossa unidade escolar na data de hoje, estando apto(a) a cursar o "
            f"<strong><u>{serie}</u></strong>."
        )
    elif tipo == "Conclusão":
        titulo = "Declaração de Conclusão"
        match = re.search(r"(\d+)º\s*ano", serie)
        if match:
            next_year = int(match.group(1)) + 1
            series_text = f"{next_year}º ano"
        else:
            series_text = "a série subsequente"
        declaracao_text = (
            f"Declaro, para os devidos fins, que o(a) aluno(a) <strong><u>{nome}</u></strong>, "
            f"portador(a) do RA <strong><u>{ra}</u></strong>, nascido(a) em <strong><u>{data_nasc}</u></strong>, "
            f"concluiu com êxito o <strong><u>{serie}</u></strong>, estando apto(a) a ingressar no "
            f"<strong><u>{series_text}</u></strong>."
        )
    else:
        titulo = "Declaração"
        declaracao_text = "Tipo de declaração inválido."
    
    base_template = f'''<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <title>{titulo} - E.M José Padin Mouta</title>
  <style>
    @page {{
      margin: 0;
    }}
    html, body {{
      margin: 0;
      padding: 0.5cm;
      font-family: 'Montserrat', sans-serif;
      font-size: 16px;
      line-height: 1.5;
      color: #333;
    }}
    .header {{
      text-align: center;
      border-bottom: 2px solid #283E51;
      padding-bottom: 5px;
      margin-bottom: 10px;
    }}
    .header h1 {{
      margin: 0;
      font-size: 24px;
      text-transform: uppercase;
      color: #283E51;
    }}
    .header p {{
      margin: 3px 0;
      font-size: 16px;
    }}
    .date {{
      text-align: right;
      font-size: 16px;
      margin-bottom: 10px;
    }}
    .content {{
      text-align: justify;
      margin-bottom: 10px;
    }}
    .signature {{
      text-align: center;
      margin: 0;
      padding: 0;
    }}
    .signature .line {{
      height: 1px;
      background-color: #333;
      width: 60%;
      margin: 0 auto 5px auto;
    }}
    .footer {{
      text-align: center;
      border-top: 2px solid #283E51;
      padding-top: 5px;
      margin: 0;
      font-size: 14px;
      color: #555;
    }}
    /* A declaração toda ficará em uma única folha */
    @media print {{
      .no-print {{ display: none !important; }}
      body {{
        margin: 0;
        padding: 0.5cm;
        font-size: 20px;
      }}
      .declaration-bottom {{
         margin-top: 10cm;
      }}

      .date {{
         margin-top: 2cm;
      }}

    }}
    {additional_css}
    /* Novo estilo para o header */
    header {{
      background: linear-gradient(90deg, #283E51, #4B79A1);
      color: #fff;
      padding: 20px;
      text-align: center;
      border-bottom: 3px solid #1d2d3a;
      border-radius: 0 0 15px 15px;
      box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }}
  </style>
</head>
<body>
  <div class="declaration-container">
    <div class="header">
      <div style="display: flex; justify-content: space-between; align-items: center;">
        <img src="/static/logos/escola.png" alt="Escola Logo" style="height: 80px;">
        <div>
          <h1>Secretaria de Educação</h1>
          <p>E.M José Padin Mouta</p>
          <p>Município da Estância Balneária de Praia Grande</p>
          <p>Estado de São Paulo</p>
        </div>
        <img src="/static/logos/municipio.png" alt="Município Logo" style="height: 80px;">
      </div>
    </div>
    <div class="date">
      <p>{data_extenso}</p>
    </div>
    <div class="content">
      <h2 style="text-align: center; text-transform: uppercase; color: #283E51;">{titulo}</h2>
      <p>{declaracao_text}</p>
    </div>
    <div class="declaration-bottom">
      <div class="signature">
        <div class="line"></div>
        <p>Luciana Rocha Augustinho</p>
        <p>Diretora da Unidade Escolar</p>
      </div>
      <div class="footer">
        <p>Rua: Bororós, nº 150, Vila Tupi, Praia Grande - SP, CEP: 11703-390</p>
        <p>Telefone: 3496-5321 | E-mail: em.padin@praiagrande.sp.gov.br</p>
      </div>
    </div>
  </div>
  <div class="no-print" style="text-align: center; margin-top: 20px;">
    <button onclick="window.print()" class="print-button">Imprimir Declaração</button>
  </div>
</body>
</html>
'''
    return base_template

@app.route('/login', methods=['GET', 'POST']) 
def login():
    error = None
    if request.method == 'POST':
        token = request.form.get('token')
        if token == ACCESS_TOKEN:
            session['logged_in'] = True
            return redirect(url_for('dashboard'))
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
            background: linear-gradient(90deg, #283E51, #4B79A1);
            color: #fff;
            padding: 20px;
            text-align: center;
            border-bottom: 3px solid #1d2d3a;
            border-radius: 0 0 15px 15px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
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
          <h1 class="mb-0">Secretaria - E.M José Padin Mouta</h1>
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

@app.route('/', methods=['GET'])
@login_required
def dashboard():
    dashboard_html = '''
    <!doctype html>
    <html lang="pt-br">
      <head>
        <meta charset="utf-8">
        <title>Secretaria - E.M José Padin Mouta</title>
        <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
        <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
        <style>
          body {
            background: #eef2f3;
            font-family: 'Montserrat', sans-serif;
          }
          header {
            background: linear-gradient(90deg, #283E51, #4B79A1);
            color: #fff;
            padding: 20px;
            text-align: center;
            border-bottom: 3px solid #1d2d3a;
            border-radius: 0 0 15px 15px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
          }
          .container-dashboard {
            background: #fff;
            padding: 40px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            margin: 40px auto;
            max-width: 800px;
          }
          .option-card {
            border: 1px solid #ccc;
            border-radius: 10px;
            padding: 20px;
            text-align: center;
            transition: transform 0.2s;
            cursor: pointer;
            margin-bottom: 20px;
          }
          .option-card:hover {
            transform: scale(1.02);
          }
          .option-card h2 {
            margin-bottom: 10px;
            color: #283E51;
          }
          .option-card p {
            color: #555;
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
        </style>
      </head>
      <body>
        <header>
          <h1>Secretaria - E.M José Padin Mouta</h1>
        </header>
        <div class="container container-dashboard">
          <div class="option-card" onclick="window.location.href='{{ url_for('declaracao_tipo') }}'">
            <h2>Declaração Escolar</h2>
            <p>Gerar declaração escolar.</p>
          </div>
          <div class="option-card" onclick="window.location.href='{{ url_for('carteirinhas') }}'">
            <h2>Carteirinhas</h2>
            <p>Gerar carteirinhas para os alunos.</p>
          </div>
          <div class="option-card" onclick="window.location.href='{{ url_for('quadros') }}'">
            <h2>Quadros</h2>
            <p>Gerar quadros para os alunos.</p>
          </div>
          <div class="logout-container">
            <a href="{{ url_for('logout') }}" class="btn-logout">
              <i class="fas fa-sign-out-alt"></i> Logout
            </a>
          </div>
        </div>
        <footer>
          Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
        </footer>
      </body>
    </html>
    '''
    return render_template_string(dashboard_html)

@app.route('/carteirinhas', methods=['GET', 'POST'])
@login_required
def carteirinhas():
    if request.method == 'POST':
        if 'excel_file' in request.files:
            file = request.files['excel_file']
            if file.filename == '':
                return "Nenhum arquivo selecionado", 400
            flash("Gerando carteirinhas. Aguarde...", "info")
            html_result = gerar_html_carteirinhas(file)
            return html_result
    carteirinhas_html = '''
    <!doctype html>
    <html lang="pt-br">
      <head>
        <meta charset="utf-8">
        <title>Carteirinhas - E.M José Padin Mouta</title>
        <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
        <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
        <style>
          body {
            background: #eef2f3;
            font-family: 'Montserrat', sans-serif;
          }
          header {
            background: linear-gradient(90deg, #283E51, #4B79A1);
            color: #fff;
            padding: 20px;
            text-align: center;
            border-bottom: 3px solid #1d2d3a;
            border-radius: 0 0 15px 15px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
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
          #flash-messages {
            position: relative;
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
            <a href="{{ url_for('logout') }}" class="btn-logout">
              <i class="fas fa-sign-out-alt"></i> Logout
            </a>
          </div>
        </div>
        <footer>
          Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
        </footer>
        <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@4.5.2/dist/js/bootstrap.bundle.min.js"></script>
        <script>
          setTimeout(function(){
            var flashDiv = document.getElementById('flash-messages');
            if(flashDiv){
              flashDiv.style.display = 'none';
            }
          }, 3000);
          
          function showLoading() {
            var existingOverlay = document.getElementById('loading-overlay');
            if (existingOverlay) {
              existingOverlay.remove();
            }
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
              <div style="text-align: center; color: white; font-family: Arial, sans-serif;">
                <svg width="3.0cm" height="4.5cm" viewBox="0 0 6.0 9.0" xmlns="http://www.w3.org/2000/svg">
                  <rect x="0.3" y="0.3" width="5.4" height="8.4" rx="0.3" ry="0.3" stroke="white" stroke-width="0.1" fill="none" />
                  <rect id="badge-fill" x="0.3" y="8.7" width="5.4" height="0" rx="0.3" ry="0.3" fill="white" />
                </svg>
                <p id="loading-text" style="margin-top: 0.2cm;">Gerando carteirinhas...</p>
              </div>
            `;
            document.body.appendChild(loadingOverlay);
            let fillHeight = 0;
            const maxHeight = 8.4;
            const interval = setInterval(() => {
              fillHeight += 0.2;
              if (fillHeight > maxHeight) {
                fillHeight = maxHeight;
                clearInterval(interval);
              }
              const badgeFill = document.getElementById('badge-fill');
              badgeFill.setAttribute('y', 8.7 - fillHeight);
              badgeFill.setAttribute('height', fillHeight);
            }, 100);
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
    return render_template_string(carteirinhas_html)

@app.route('/declaracao/upload', methods=['GET', 'POST'])
@login_required
def declaracao_upload():
    if request.method == 'POST':
        if 'excel_file' not in request.files:
            flash("Nenhum arquivo enviado.", "error")
            return redirect(url_for('declaracao_upload'))
        file = request.files['excel_file']
        if file.filename == '':
            flash("Nenhum arquivo selecionado.", "error")
            return redirect(url_for('declaracao_upload'))
        filename = secure_filename(file.filename)
        unique_filename = f"declaracao_{uuid.uuid4().hex}_{filename}"
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
        file.save(file_path)
        session['declaracao_excel'] = file_path
        return redirect(url_for('declaracao_select'))
    
    upload_form = '''
    <!doctype html>
    <html lang="pt-br">
    <head>
      <meta charset="utf-8">
      <title>Declaração Escolar - Anexar Lista Piloto</title>
      <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
      <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
      <style>
        body {
          background: #eef2f3;
          font-family: 'Montserrat', sans-serif;
        }
        header {
          background: linear-gradient(90deg, #283E51, #4B79A1);
          color: #fff;
          padding: 20px;
          text-align: center;
          border-bottom: 3px solid #1d2d3a;
          border-radius: 0 0 15px 15px;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .container-form {
          background: #fff;
          padding: 40px;
          border-radius: 10px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.15);
          margin: 40px auto;
          max-width: 600px;
          text-align: center;
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
          position: fixed;
          bottom: 0;
          width: 100%;
        }
      </style>
    </head>
    <body>
      <header>
        <h1>Declaração Escolar - Anexar Lista Piloto</h1>
      </header>
      <div class="container container-form">
        <form method="POST" enctype="multipart/form-data">
          <div class="form-group">
            <label for="excel_file">Selecione o arquivo Excel:</label>
            <input type="file" class="form-control-file" name="excel_file" id="excel_file" accept=".xlsx, .xls" required>
          </div>
          <button type="submit" class="btn btn-primary">Anexar Lista</button>
        </form>
      </div>
      <footer>
        Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
      </footer>
    </body>
    </html>
    '''
    return render_template_string(upload_form)

@app.route('/declaracao/select', methods=['GET', 'POST'])
@login_required
def declaracao_select():
    file_path = session.get('declaracao_excel')
    if not file_path or not os.path.exists(file_path):
        flash("Arquivo Excel não encontrado. Por favor, anexe a lista piloto.", "error")
        return redirect(url_for('declaracao_upload'))
    
    planilha = pd.read_excel(file_path, sheet_name='LISTA CORRIDA')
    def format_rm(x):
        try:
            return str(int(float(x)))
        except:
            return str(x)
    planilha['RM_str'] = planilha['RM'].apply(format_rm)
    alunos = planilha[planilha['RM_str'] != "0"][['RM_str', 'NOME']].drop_duplicates()
    options_html = ""
    for _, row in alunos.iterrows():
        rm_str = row['RM_str']
        nome = row['NOME']
        options_html += f'<option value="{rm_str}">{rm_str} - {nome}</option>'
    
    if request.method == 'POST':
        rm = request.form.get('rm')
        tipo = request.form.get('tipo')
        if not rm or not tipo:
            flash("Selecione o aluno e o tipo de declaração.", "error")
            return redirect(url_for('declaracao_select'))
        declaracao_html = gerar_declaracao_escolar(file_path, rm, tipo)
        if declaracao_html is None:
            flash("Aluno não encontrado.", "error")
            return redirect(url_for('declaracao_select'))
        return declaracao_html

    select_form = f'''
    <!doctype html>
    <html lang="pt-br">
    <head>
      <meta charset="utf-8">
      <title>Declaração Escolar - Seleção de Aluno</title>
      <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
      <link href="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/css/select2.min.css" rel="stylesheet" />
      <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
      <style>
        body {{
          background: #eef2f3;
          font-family: 'Montserrat', sans-serif;
        }}
        header {{
          background: linear-gradient(90deg, #283E51, #4B79A1);
          color: #fff;
          padding: 20px;
          text-align: center;
          border-bottom: 3px solid #1d2d3a;
          border-radius: 0 0 15px 15px;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        .container-form {{
          background: #fff;
          padding: 40px;
          border-radius: 10px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.15);
          margin: 40px auto;
          max-width: 600px;
        }}
        .btn-primary {{
          background-color: #283E51;
          border: none;
        }}
        .btn-primary:hover {{
          background-color: #1d2d3a;
        }}
        footer {{
          background-color: #424242;
          color: #fff;
          text-align: center;
          padding: 10px;
          position: fixed;
          bottom: 0;
          width: 100%;
        }}
      </style>
    </head>
    <body>
      <header>
        <h1>Declaração Escolar</h1>
        <p>Selecione o aluno e o tipo de declaração</p>
      </header>
      <div class="container container-form">
        <form method="POST">
          <div class="form-group">
            <label for="rm">Aluno:</label>
            <select class="form-control" id="rm" name="rm" required>
              <option value="">Selecione</option>
              {options_html}
            </select>
          </div>
          <div class="form-group">
            <label for="tipo">Tipo de Declaração:</label>
            <select class="form-control" id="tipo" name="tipo" required>
              <option value="">Selecione</option>
              <option value="Escolaridade">Declaração de Escolaridade</option>
              <option value="Transferencia">Declaração de Transferência</option>
              <option value="Conclusão">Declaração de Conclusão</option>
            </select>
          </div>
          <button type="submit" class="btn btn-primary">Gerar Declaração</button>
        </form>
      </div>
      <footer>
        Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
      </footer>
      <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
      <script src="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/js/select2.min.js"></script>
      <script>
        $(document).ready(function() {{
          $('#rm').select2({{
            placeholder: "Selecione o aluno",
            allowClear: true
          }});
        }});
      </script>
    </body>
    </html>
    '''
    return render_template_string(select_form)

@app.route('/declaracao/tipo', methods=['GET', 'POST'])
@login_required
def declaracao_tipo():
    if request.method == 'POST':
        tipo = request.form.get('tipo')
        if tipo == 'Fundamental':
            return redirect(url_for('declaracao_upload'))
        elif tipo == 'EJA':
            msg_html = '''
            <!doctype html>
            <html lang="pt-br">
            <head>
                 <meta charset="utf-8">
                 <title>Declaração EJA - Em Desenvolvimento</title>
                 <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
                 <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
                 <style>
                 body {
                     background: #eef2f3;
                     font-family: 'Montserrat', sans-serif;
                 }
                 header {
                     background: linear-gradient(90deg, #283E51, #4B79A1);
                     color: #fff;
                     padding: 20px;
                     text-align: center;
                     border-bottom: 3px solid #1d2d3a;
                     border-radius: 0 0 15px 15px;
                     box-shadow: 0 4px 6px rgba(0,0,0,0.1);
                 }
                 .container-msg {
                     background: #fff;
                     padding: 40px;
                     border-radius: 10px;
                     box-shadow: 0 4px 12px rgba(0,0,0,0.15);
                     margin: 40px auto;
                     max-width: 600px;
                     text-align: center;
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
                 </style>
            </head>
            <body>
                 <header>
                     <h1>Declaração EJA</h1>
                 </header>
                 <div class="container-msg">
                     <p>Declaração EJA está em desenvolvimento.</p>
                     <a href="{{ url_for('dashboard') }}" class="btn btn-primary">Voltar ao Dashboard</a>
                 </div>
                 <footer>
                     Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
                 </footer>
            </body>
            </html>
            '''
            return render_template_string(msg_html)
    form_html = '''
    <!doctype html>
    <html lang="pt-br">
    <head>
         <meta charset="utf-8">
         <title>Selecionar Tipo de Declaração Escolar</title>
         <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
         <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
         <style>
         body {
             background: #eef2f3;
             font-family: 'Montserrat', sans-serif;
         }
         header {
             background: linear-gradient(90deg, #283E51, #4B79A1);
             color: #fff;
             padding: 20px;
             text-align: center;
             border-bottom: 3px solid #1d2d3a;
             border-radius: 0 0 15px 15px;
             box-shadow: 0 4px 6px rgba(0,0,0,0.1);
         }
         .container-form {
             background: #fff;
             padding: 40px;
             border-radius: 10px;
             box-shadow: 0 4px 12px rgba(0,0,0,0.15);
             margin: 40px auto;
             max-width: 600px;
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
             position: fixed;
             bottom: 0;
             width: 100%;
         }
         </style>
    </head>
    <body>
         <header>
             <h1>Selecionar Tipo de Declaração Escolar</h1>
         </header>
         <div class="container-form">
             <form method="POST">
                 <div class="form-group">
                     <label for="tipo">Selecione o tipo:</label>
                     <select class="form-control" id="tipo" name="tipo" required>
                         <option value="">Selecione</option>
                         <option value="Fundamental">Declaração Fundamental</option>
                         <option value="EJA">Declaração EJA</option>
                     </select>
                 </div>
                 <button type="submit" class="btn btn-primary">Continuar</button>
             </form>
         </div>
         <footer>
             Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
         </footer>
    </body>
    </html>
    '''
    return render_template_string(form_html)

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
    return redirect(url_for('carteirinhas'))

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
    return redirect(url_for('carteirinhas'))

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

@app.route('/quadros')
@login_required
def quadros():
    quadros_html = '''
    <!doctype html>
    <html lang="pt-br">
    <head>
      <meta charset="utf-8">
      <title>Quadros - E.M José Padin Mouta</title>
      <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
      <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
      <style>
        body { background: #eef2f3; font-family: 'Montserrat', sans-serif; }
        header {
          background: linear-gradient(90deg, #283E51, #4B79A1);
          color: #fff;
          padding: 20px;
          text-align: center;
          border-bottom: 3px solid #1d2d3a;
          border-radius: 0 0 15px 15px;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .container-menu { 
          margin: 40px auto; 
          max-width: 800px; 
          background: #fff; 
          padding: 40px; 
          border-radius: 10px; 
          box-shadow: 0 4px 12px rgba(0,0,0,0.15); 
        }
        .option-card { border: 1px solid #ccc; border-radius: 10px; padding: 20px; text-align: center; transition: transform 0.2s; cursor: pointer; margin-bottom: 20px; }
        .option-card:hover { transform: scale(1.02); }
        .option-card h2 { margin-bottom: 10px; color: #283E51; }
        .option-card p { color: #555; }
        footer { background-color: #424242; color: #fff; text-align: center; padding: 10px; position: fixed; bottom: 0; width: 100%; }
      </style>
    </head>
    <body>
      <header>
        <h1>Quadros - E.M José Padin Mouta</h1>
      </header>
      <div class="container-menu">
        <div class="option-card" onclick="window.location.href='{{ url_for('quadros_inclusao') }}'">
          <h2>Inclusão</h2>
          <p>Gerar quadro de inclusão.</p>
        </div>
        <div class="option-card" onclick="window.location.href='{{ url_for('quadros_quantitativo') }}'">
          <h2>Quantitativo</h2>
          <p>Gerar quadro quantitativo.</p>
        </div>
        <div class="option-card" onclick="alert('Funcionalidade Transferências em desenvolvimento')">
          <h2>Transferências</h2>
          <p>Gerar quadro de transferências.</p>
        </div>
        <div class="option-card" onclick="window.location.href='{{ url_for('dashboard') }}'">
          <h2>Voltar ao Dashboard</h2>
        </div>
      </div>
      <footer>
        Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
      </footer>
    </body>
    </html>
    '''
    return render_template_string(quadros_html)

# Rota para Quadro de Inclusão (já existente)
@app.route('/quadros/inclusao', methods=['GET', 'POST'])
@login_required
def quadros_inclusao():
    if request.method == 'POST':
        if 'modelo_file' not in request.files or 'lista_file' not in request.files:
            flash("Arquivos não enviados.", "error")
            return redirect(url_for('quadros_inclusao'))
        modelo_file = request.files['modelo_file']
        lista_file = request.files['lista_file']
        if modelo_file.filename == '' or lista_file.filename == '':
            flash("Selecione os dois arquivos.", "error")
            return redirect(url_for('quadros_inclusao'))
        try:
            modelo_file.seek(0)
            wb = load_workbook_model(modelo_file)
        except Exception as e:
            flash(f"Erro ao ler o arquivo de modelo: {str(e)}", "error")
            return redirect(url_for('quadros_inclusao'))
        ws = wb.active
        set_merged_cell_value(ws, "C3", "Luciana Rocha Augustinho")
        set_merged_cell_value(ws, "H3", "Ana Carolina Valencio da Silva Rodrigues")
        set_merged_cell_value(ws, "K3", "Rosemeire de Souza Pereira")
        set_merged_cell_value(ws, "C4", "Rafael Marques Lima")
        set_merged_cell_value(ws, "H4", "Rita de Cassia de Andrade")
        set_merged_cell_value(ws, "K4", "Ana Paula Rodrigues de Assis Santos")
        set_merged_cell_value(ws, "P4", datetime.now().strftime("%d/%m/%Y"))
        try:
            lista_file.seek(0)
            df = pd.read_excel(lista_file, sheet_name="LISTA CORRIDA")
        except Exception as e:
            flash("Erro ao ler a Lista Piloto.", "error")
            return redirect(url_for('quadros_inclusao'))
        if len(df.columns) < 16:
            flash("O arquivo da Lista Piloto não possui colunas suficientes.", "error")
            return redirect(url_for('quadros_inclusao'))
        inclusion_col = df.columns[13]
        start_row = 7
        current_row = start_row
        for idx, row in df.iterrows():
            if not str(row['RM']).strip() or str(row['RM']).strip() == "0":
                continue
            if str(row[inclusion_col]).strip().lower() == "sim":
                col_a_val = str(row[df.columns[0]]).strip()
                match = re.match(r"(\d+º).*?([A-Za-z])$", col_a_val)
                if match:
                    nivel = match.group(1)
                    turma = match.group(2)
                else:
                    nivel = col_a_val
                    turma = ""
                horario = str(row[df.columns[10]]).strip()
                if "08h" in horario and "12h" in horario:
                    periodo = "MANHÃ"
                elif horario == "13h30 às 17h30":
                    periodo = "TARDE"
                elif horario == "19h00 às 23h00":
                    periodo = "NOITE"
                else:
                    periodo = ""
                nome_aluno = str(row[df.columns[3]]).strip()
                data_nasc = row[df.columns[5]]
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
                professor = str(row[df.columns[14]]).strip()
                plano = str(row[df.columns[15]]).strip()
                aee = str(row[df.columns[16]]).strip() if len(df.columns) > 16 else ""
                deficiencia = str(row[df.columns[17]]).strip() if len(df.columns) > 17 else ""
                observacoes = str(row[df.columns[18]]).strip() if len(df.columns) > 18 else ""
                cadeira = str(row[df.columns[19]]).strip() if len(df.columns) > 19 else ""
                adequacoes = cadeira
                atendimentos = "-"
                ws.cell(row=current_row, column=2, value=nivel)
                ws.cell(row=current_row, column=3, value=turma)
                ws.cell(row=current_row, column=4, value=periodo)
                ws.cell(row=current_row, column=5, value=horario)
                ws.cell(row=current_row, column=6, value=nome_aluno)
                ws.cell(row=current_row, column=7, value=data_nasc)
                ws.cell(row=current_row, column=8, value=professor)
                ws.cell(row=current_row, column=9, value=plano)
                ws.cell(row=current_row, column=10, value=aee)
                ws.cell(row=current_row, column=11, value=deficiencia)
                ws.cell(row=current_row, column=12, value=observacoes)
                ws.cell(row=current_row, column=13, value=cadeira)
                ws.cell(row=current_row, column=14, value=adequacoes)
                ws.cell(row=current_row, column=15, value=atendimentos)
                current_row += 1
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        filename = f"Quadro_Inclusao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    upload_html = '''
    <!doctype html>
    <html lang="pt-br">
    <head>
      <meta charset="utf-8">
      <title>Quadro de Inclusão - E.M José Padin Mouta</title>
      <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
      <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
      <style>
        body { background: #eef2f3; font-family: 'Montserrat', sans-serif; }
        header {
          background: linear-gradient(90deg, #283E51, #4B79A1);
          color: #fff;
          padding: 20px;
          text-align: center;
          border-bottom: 3px solid #1d2d3a;
          border-radius: 0 0 15px 15px;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .container-form { background: #fff; padding: 40px; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); margin: 40px auto; max-width: 600px; }
        .btn-primary { background-color: #283E51; border: none; }
        .btn-primary:hover { background-color: #1d2d3a; }
        footer { background-color: #424242; color: #fff; text-align: center; padding: 10px; position: fixed; bottom: 0; width: 100%; }
      </style>
    </head>
    <body>
      <header>
        <h1>Quadro de Inclusão</h1>
      </header>
      <div class="container-form">
        <form method="POST" enctype="multipart/form-data">
          <div class="form-group">
            <label for="modelo_file">Selecione o Modelo de Quadro (Excel):</label>
            <input type="file" class="form-control-file" name="modelo_file" id="modelo_file" accept=".xlsx, .xls" required>
          </div>
          <div class="form-group">
            <label for="lista_file">Selecione a Lista Piloto (Excel):</label>
            <input type="file" class="form-control-file" name="lista_file" id="lista_file" accept=".xlsx, .xls" required>
          </div>
          <button type="submit" class="btn btn-primary">Gerar Quadro de Inclusão</button>
        </form>
        <br>
        <a href="{{ url_for('quadros') }}">Voltar para Quadros</a>
      </div>
      <footer>
        Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
      </footer>
    </body>
    </html>
    '''
    return render_template_string(upload_html)

# Rota para Quadro Quantitativo – agora preenchido com dados da aba "Total de Alunos"
@app.route('/quadros/quantitativo', methods=['GET', 'POST'])
@login_required
def quadros_quantitativo():
    if request.method == 'POST':
        if 'modelo_file' not in request.files or 'lista_file' not in request.files:
            flash("Arquivos não enviados.", "error")
            return redirect(url_for('quadros_quantitativo'))
        modelo_file = request.files['modelo_file']
        lista_file = request.files['lista_file']
        if modelo_file.filename == '' or lista_file.filename == '':
            flash("Selecione os dois arquivos.", "error")
            return redirect(url_for('quadros_quantitativo'))
        
        # Para o arquivo modelo, aceitamos XLS ou XLSX
        try:
            wb_modelo = load_workbook_model(modelo_file)
        except Exception as e:
            flash(f"Erro ao ler o arquivo de modelo: {str(e)}", "error")
            return redirect(url_for('quadros_quantitativo'))
        
        # Preencher na segunda aba (Aba 2: Fundamental..EJA) se existir; caso contrário, usa a aba ativa
        if len(wb_modelo.sheetnames) >= 2:
            ws_modelo = wb_modelo.worksheets[1]
        else:
            ws_modelo = wb_modelo.active

        # Preencher o modelo conforme especificado (usando set_merged_cell_value para tratar células mescladas)
        set_merged_cell_value(ws_modelo, "B5", "E.M José Padin Mouta")
        set_merged_cell_value(ws_modelo, "C6", "Rafael Fernando da Silva")
        set_merged_cell_value(ws_modelo, "B7", "46034")
        current_month = datetime.now().strftime("%m")
        set_merged_cell_value(ws_modelo, "A13", f"{current_month}/2025")
        try:
            lista_file.seek(0)
            wb_lista = load_workbook(lista_file, data_only=True)
        except Exception as e:
            flash("Erro ao ler o arquivo da lista piloto.", "error")
            return redirect(url_for('quadros_quantitativo'))
        # Busca a aba "Total de Alunos" de forma case-insensitive
        sheet_name = None
        for name in wb_lista.sheetnames:
            if name.strip().lower() == "total de alunos":
                sheet_name = name
                break
        if not sheet_name:
            flash("A aba 'Total de Alunos' não foi encontrada na lista piloto.", "error")
            return redirect(url_for('quadros_quantitativo'))
        ws_total = wb_lista[sheet_name]
        # Preencher linhas 37 a 42 utilizando os valores da aba "Total de Alunos"
        for r in range(37, 43):
            source_row = r - 31  # linhas 6 a 11
            value_B = ws_total.cell(row=source_row, column=7).value  # coluna G
            value_C = ws_total.cell(row=source_row, column=8).value  # coluna H
            set_merged_cell_value(ws_modelo, f"B{r}", value_B)
            set_merged_cell_value(ws_modelo, f"C{r}", value_C)
            set_merged_cell_value(ws_modelo, f"D{r}", f"=B{r}+C{r}")
        output = BytesIO()
        wb_modelo.save(output)
        output.seek(0)
        filename = f"Quadro_Quantitativo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    upload_html = '''
    <!doctype html>
    <html lang="pt-br">
    <head>
      <meta charset="utf-8">
      <title>Quadro Quantitativo - E.M José Padin Mouta</title>
      <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
      <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
      <style>
        body { background: #eef2f3; font-family: 'Montserrat', sans-serif; }
        header {
          background: linear-gradient(90deg, #283E51, #4B79A1);
          color: #fff;
          padding: 20px;
          text-align: center;
          border-bottom: 3px solid #1d2d3a;
          border-radius: 0 0 15px 15px;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .container-form { background: #fff; padding: 40px; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); margin: 40px auto; max-width: 600px; }
        .btn-primary { background-color: #283E51; border: none; }
        .btn-primary:hover { background-color: #1d2d3a; }
        footer { background-color: #424242; color: #fff; text-align: center; padding: 10px; position: fixed; bottom: 0; width: 100%; }
      </style>
    </head>
    <body>
      <header>
        <h1>Quadro Quantitativo</h1>
      </header>
      <div class="container-form">
        <form method="POST" enctype="multipart/form-data">
          <div class="form-group">
            <label for="modelo_file">Selecione o Modelo de Quadro (Excel):</label>
            <input type="file" class="form-control-file" name="modelo_file" id="modelo_file" accept=".xlsx, .xls" required>
          </div>
          <div class="form-group">
            <label for="lista_file">Selecione a Lista Piloto (Excel):</label>
            <input type="file" class="form-control-file" name="lista_file" id="lista_file" accept=".xlsx, .xls" required>
          </div>
          <button type="submit" class="btn btn-primary">Gerar Quadro Quantitativo</button>
        </form>
        <br>
        <a href="{{ url_for('quadros') }}">Voltar para Quadros</a>
      </div>
      <footer>
        Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
      </footer>
    </body>
    </html>
    '''
    return render_template_string(upload_html)

if __name__ == '__main__':
    app.run(debug=True)
