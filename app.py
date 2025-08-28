from flask import Flask, request
import gspread
from google.oauth2.service_account import Credentials
import time
import os

app = Flask(__name__)

# ---------------- CONFIGURAÇÃO ----------------
# Nome do arquivo JSON de credenciais (coloque na mesma pasta do script)
CREDS_FILE = "third-hope-421922-0ff458bfa5d2.json"

# ID da planilha (fornecido por você)
SHEET_ID = "18fPXEnrApXVaEV1PTu3_qKaFdFAfhAh1UVnqh_AHbcU"

# Escopo necessário para ler e escrever na planilha
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Timeout ao aguardar o cálculo (segundos)
RESULT_TIMEOUT = 6.0
POLL_INTERVAL = 0.5
# ------------------------------------------------

# Valida se o arquivo de credenciais existe
if not os.path.exists(CREDS_FILE):
    raise FileNotFoundError(
        f"Arquivo de credenciais '{CREDS_FILE}' não encontrado. "
        "Coloque o JSON da conta de serviço no mesmo diretório e atualize CREDS_FILE."
    )

# Autentica e abre a planilha
creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
client = gspread.authorize(creds)
sheet = client.open_by_key(SHEET_ID).sheet1  # primeira aba

# HTML do formulário (mantive seu estilo)
FORM_HTML = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>Orçamento Porta de Guilhotina</title>
<style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap');

    body {
        font-family: 'Roboto', sans-serif;
        background: linear-gradient(135deg, #00c6ff, #0072ff);
        margin: 0;
        padding: 0;
        display: flex;
        justify-content: center;
        align-items: center;
        min-height: 100vh;
    }

    .container {
        background-color: #ffffff;
        border-radius: 12px;
        padding: 40px;
        box-shadow: 0 8px 25px rgba(0,0,0,0.2);
        max-width: 500px;
        width: 100%;
        text-align: center;
    }

    h2 {
        color: #0072ff;
        margin-bottom: 30px;
    }

    form label {
        display: block;
        text-align: left;
        margin-bottom: 6px;
        font-weight: 500;
        color: #333;
    }

    form input[type="text"], form select {
        width: 100%;
        padding: 10px 15px;
        margin-bottom: 20px;
        border-radius: 8px;
        border: 1px solid #ccc;
        font-size: 16px;
        transition: all 0.3s;
    }

    form input[type="text"]:focus, form select:focus {
        border-color: #0072ff;
        box-shadow: 0 0 5px rgba(0,114,255,0.5);
        outline: none;
    }

    button {
        background-color: #0072ff;
        color: white;
        border: none;
        padding: 12px 25px;
        font-size: 16px;
        border-radius: 8px;
        cursor: pointer;
        transition: background 0.3s;
    }

    button:hover {
        background-color: #005bbb;
    }

    .resultado {
        margin-top: 20px;
        font-size: 20px;
        font-weight: 700;
        color: red; /* destaque vermelho */
    }

    .erro {
        margin-top: 12px;
        font-size: 14px;
        color: darkred;
        font-weight: 600;
    }
</style>
</head>
<body>
<div class="container">
<h2>Formulário de Orçamento</h2>
<form method="POST" action="/calcular">
    <label>Orçamento:</label>
    <input type="text" name="orcamento" required>

    <label>Nome do Cliente:</label>
    <input type="text" name="cliente" required>

    <label>Tipo da Porta:</label>
    <select name="porta" required>
        <option value="Simples">Simples</option>
        <option value="Em L Esquerda">Em L Esquerda</option>
        <option value="Em L Direita">Em L Direita</option>
        <option value="Em U">Em U</option>
    </select>

    <label>Altura Total (cm):</label>
    <input type="text" name="altura" required>

    <label>Largura da Porta (cm):</label>
    <input type="text" name="largura" required>

    <label>Profundidade Total (cm):</label>
    <input type="text" name="profundidade" required>

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
<div id="erro" class="erro"></div>
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

    # Proteção: garante que recebemos as chaves esperadas
    expected = ["orcamento", "cliente", "porta", "altura", "largura", "profundidade", "pedra", "vidro"]
    for k in expected:
        if k not in dados:
            return FORM_HTML.replace(
                '<div id="erro" class="erro"></div>',
                f'<div id="erro" class="erro">Campo {k} ausente no formulário.</div>'
            )

    try:
        # Atualiza as células (pode-se usar batch_update, mas mantive simples)
        sheet.update("B2", dados["orcamento"])
        sheet.update("B3", dados["cliente"])
        sheet.update("B4", dados["porta"])
        sheet.update("B5", dados["altura"])
        sheet.update("B6", dados["largura"])
        sheet.update("B7", dados["profundidade"])
        sheet.update("B10", dados["pedra"])
        sheet.update("B12", dados["vidro"])
    except Exception as e:
        # Erro ao escrever na planilha
        return FORM_HTML.replace(
            '<div id="erro" class="erro"></div>',
            f'<div id="erro" class="erro">Erro ao escrever na planilha: {str(e)}</div>'
        )

    # Aguarda (poll) até que a célula D10 tenha um valor (timeout)
    inicio = time.time()
    resultado = None
    while (time.time() - inicio) < RESULT_TIMEOUT:
        try:
            resultado = sheet.acell("D10").value
            # Se encontrou algo não vazio, sai
            if resultado not in (None, "", " "):
                break
        except Exception:
            # leitura pode falhar pontualmente, ignora e tenta novamente
            resultado = None
        time.sleep(POLL_INTERVAL)

    if resultado in (None, "", " "):
        # Se não obteve resultado no tempo limite, retorna aviso
        return FORM_HTML.replace(
            '<div id="erro" class="erro"></div>',
            '<div id="erro" class="erro">Resultado indisponível. Tente novamente em alguns segundos.</div>'
        )

    # Converte e formata o resultado para exibição (tenta tratar vírgula/ponto)
    try:
        # resultado pode vir "1234.56" ou "1.234,56" etc -> normaliza
        r_str = str(resultado).strip()
        r_str = r_str.replace(".", "").replace(",", ".") if r_str.count(",") == 1 and r_str.count(".") > 0 else r_str
        # fallback: substituir virgula por ponto se houver
        r_str = r_str.replace(",", ".")
        valor = float(r_str)
        resultado_formatado = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        # Se não for convertível, mostra o valor bruto
        resultado_formatado = str(resultado)

    return FORM_HTML.replace(
        '<div id="resultado" class="resultado"></div>',
        f'<div id="resultado" class="resultado">Custo de Produção: R$ {resultado_formatado}</div>'
    )

if __name__ == "__main__":
    # Em produção use gunicorn: gunicorn app:app
    app.run(host="0.0.0.0", port=5000, debug=True)
