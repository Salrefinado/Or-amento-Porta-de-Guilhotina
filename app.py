from flask import Flask, request
import xlwings as xw

app = Flask(__name__)

CAMINHO = r"C:\Users\User\Desktop\Bunese\Porta de Guilhotina Site.xlsm"

# HTML do formulário estilizado
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

    wb = xw.Book(CAMINHO)
    sheet = wb.sheets[0]

    # Preenche apenas as células correspondentes aos campos existentes no formulário
    sheet.range("B2").value = dados["orcamento"]
    sheet.range("B3").value = dados["cliente"]
    sheet.range("B4").value = dados["porta"]
    sheet.range("B5").value = dados["altura"]
    sheet.range("B6").value = dados["largura"]
    sheet.range("B7").value = dados["profundidade"]
    sheet.range("B10").value = dados["pedra"]
    sheet.range("B12").value = dados["vidro"]

    wb.app.calculate()
    resultado = sheet.range("D10").value

    # Formata o valor com 2 casas decimais e vírgula
    resultado_formatado = f"{resultado:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    # Retorna o HTML atualizado com o custo em vermelho
    return FORM_HTML.replace(
        '<div id="resultado" class="resultado"></div>',
        f'<div id="resultado" class="resultado">Custo de Produção: R$ {resultado_formatado}</div>'
    )

if __name__ == "__main__":
    app.run(debug=True)
