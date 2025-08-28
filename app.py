from flask import Flask, request
import gspread
from oauth2client.service_account import ServiceAccountCredentials

app = Flask(__name__)

# Caminho para o JSON da conta de serviço
CREDENCIAIS = "credenciais.json"
# Nome da planilha
PLANILHA = "Porta de Guilhotina"

# Autenticação Google Sheets
scope = ["https://spreadsheets.google.com/feeds","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENCIAIS, scope)
client = gspread.authorize(creds)
sheet = client.open(PLANILHA).sheet1

# HTML do formulário
FORM_HTML = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Orçamento Porta de Guilhotina</title>
<style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap');
    body { font-family: 'Roboto', sans-serif; background: linear-gradient(135deg, #00c6ff, #0072ff); display: flex; justify-content: center; align-items: center; min-height: 100vh; margin:0; }
    .container { background-color:#fff; border-radius:12px; padding:40px; box-shadow:0 8px 25px rgba(0,0,0,0.2); max-width:500px; width:100%; text-align:center; }
    h2 { color:#0072ff; margin-bottom:30px; }
    form label { display:block; text-align:left; margin-bottom:6px; font-weight:500; color:#333; }
    form input[type="text"], form select { width:100%; padding:10px 15px; margin-bottom:20px; border-radius:8px; border:1px solid #ccc; font-size:16px; transition: all 0.3s; }
    form input[type="text"]:focus, form select:focus { border-color:#0072ff; box-shadow:0 0 5px rgba(0,114,255,0.5); outline:none; }
    button { background-color:#0072ff; color:white; border:none; padding:12px 25px; font-size:16px; border-radius:8px; cursor:pointer; transition: background 0.3s; }
    button:hover { background-color:#005bbb; }
    .resultado { margin-top:20px; font-size:20px; font-weight:700; color:red; }
</style>
</head>
<body>
<div class="container">
<h2>Formulário de Orçamento</h2>
<form method="POST" action="/calcular">
    <label>Orçamento:</label><input type="text" name="orcamento" required>
    <label>Nome do Cliente:</label><input type="text" name="cliente" required>
    <label>Tipo da Porta:</label>
    <select name="porta" required>
        <option value="Simples">Simples</option>
        <option value="Em L Esquerda">Em L Esquerda</option>
        <option value="Em L Direita">Em L Direita</option>
        <option value="Em U">Em U</option>
    </select>
    <label>Altura Total (cm):</label><input type="text" name="altura" required>
    <label>Largura da Porta (cm):</label><input type="text" name="largura" required>
    <label>Profundidade Total (cm):</label><input type="text" name="profundidade" required>
    <label>Pedra instalada:</label>
    <select name="pedra" required>
        <option value="Não">Não</option>
        <option value="Sim">Sim</option>
    </select>
    <label>Fechamento em Vidro:</label>
    <select name="vidro" required>
        <option value="Sim">Sim</option>
        <option value="Não">Não</option>
    </select>
    <button type="submit">Calcular Custo</button>
</form>
<div id="resultado" class="resultado"></div>
</div>
</body>
</html>
"""

@app.route("/", methods=["GET"])
def index():
    return FORM_HTML

@app.route("/calcular", methods=["POST"])
def calcular():
    dados = request.form

    # Atualiza células no Google Sheets
    sheet.update("B2", dados["orcamento"])
    sheet.update("B3", dados["cliente"])
    sheet.update("B4", dados["porta"])
    sheet.update("B5", dados["altura"])
    sheet.update("B6", dados["largura"])
    sheet.update("B7", dados["profundidade"])
    sheet.update("B10", dados["pedra"])
    sheet.update("B12", dados["vidro"])

    # Lê resultado
    resultado = sheet.acell("D10").value

    # Formata valor
    try:
        valor = float(resultado)
        resultado_formatado = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        resultado_formatado = resultado

    return FORM_HTML.replace(
        '<div id="resultado" class="resultado"></div>',
        f'<div id="resultado" class="resultado">Custo de Produção: R$ {resultado_formatado}</div>'
    )

if __name__ == "__main__":
    app.run(debug=True)
