from flask import Flask, render_template, request
from datetime import datetime, timedelta
import os
from werkzeug.utils import secure_filename

from docxtpl import DocxTemplate

import smtplib
from email.message import EmailMessage
import mimetypes

# ================= CONFIGURAÇÕES BÁSICAS =================

app = Flask(__name__)

BASE_DIR = os.path.abspath(os.path.dirname(__file__))

TEMPLATE_DOCX = os.path.join(BASE_DIR, "contrato_template.docx")
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "outputs")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# --------- CARROS CADASTRADOS ---------
CARROS = {
    "prisma_2019": {
        "nome_exibicao": "Prisma 2019 – Prata – PLM7A56",
        "marca": "CHEVROLET",
        "modelo": "PRISMA",
        "ano": "2019",
        "cor": "PRATA",
        "placa": "PLM7A56",
        "categoria": "SEDAN",
        "valor_avaliacao": "50.000,00",
    },
    "polo_2026": {
        "nome_exibicao": "Polo 2026 – Prata – RSR9I01",
        "marca": "VOLKSWAGEN",
        "modelo": "POLO",
        "ano": "2026",
        "cor": "PRATA",
        "placa": "RSR9I01",
        "categoria": "HATCH",
        "valor_avaliacao": "90.000,00",
    },
    "gol_2017": {
        "nome_exibicao": "Gol 2017 – Branca – PSW9J70",
        "marca": "VOLKSWAGEN",
        "modelo": "GOL",
        "ano": "2017",
        "cor": "BRANCA",
        "placa": "PSW9J70",
        "categoria": "PARTICULAR",
        "valor_avaliacao": "45.000,00",
    },
}

# ================= CONFIG E-MAIL (GMAIL + VARIÁVEIS DE AMBIENTE) =================

SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

SMTP_USER = os.environ.get("SMTP_USER")          # ex: seuemail@gmail.com
SMTP_PASSWORD = os.environ.get("SMTP_PASSWORD")  # senha de app do Gmail
EMAIL_DESTINO = os.environ.get("EMAIL_DESTINO", SMTP_USER)


# ================= FUNÇÕES AUXILIARES =================

def up(s: str) -> str:
    """Converte texto para MAIÚSCULO (se não for None)."""
    return s.upper() if s else s


def salvar_upload(file_obj):
    """Salva arquivo enviado e retorna o caminho completo."""
    if not file_obj or file_obj.filename == "":
        return None
    nome_limpo = secure_filename(file_obj.filename)
    nome_final = datetime.now().strftime("%Y%m%d%H%M%S_") + nome_limpo
    caminho = os.path.join(UPLOAD_FOLDER, nome_final)
    file_obj.save(caminho)
    return caminho


def enviar_email_com_anexos(assunto, corpo, destinatario, caminhos):
    """Envia e-mail com 1 ou mais anexos."""
    msg = EmailMessage()
    msg["Subject"] = assunto
    msg["From"] = SMTP_USER
    msg["To"] = destinatario
    msg.set_content(corpo)

    for caminho in caminhos:
        if not caminho:
            continue
        if not os.path.exists(caminho):
            continue

        mime_type, _ = mimetypes.guess_type(caminho)
        if mime_type:
            maintype, subtype = mime_type.split("/", 1)
        else:
            maintype, subtype = "application", "octet-stream"

        with open(caminho, "rb") as f:
            dados = f.read()

        nome_arquivo = os.path.basename(caminho)

        msg.add_attachment(
            dados,
            maintype=maintype,
            subtype=subtype,
            filename=nome_arquivo,
        )

    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.starttls()
        smtp.login(SMTP_USER, SMTP_PASSWORD)
        smtp.send_message(msg)


# ================= ROTAS =================

@app.route("/", methods=["GET", "POST"])
def formulario():
    if request.method == "GET":
        # Mostra formulário para o cliente
        return render_template("form.html", carros=CARROS)

    # ---------- 1. RECEBE OS DADOS DO FORMULÁRIO ----------
    locatario_nome = request.form.get("locatario_nome")
    locatario_nacionalidade = request.form.get("locatario_nacionalidade")
    locatario_estado_civil = request.form.get("locatario_estado_civil")
    locatario_profissao = request.form.get("locatario_profissao")
    locatario_rg = request.form.get("locatario_rg")
    locatario_cpf = request.form.get("locatario_cpf")
    locatario_cnh = request.form.get("locatario_cnh")

    locatario_rua = request.form.get("locatario_rua")
    locatario_numero = request.form.get("locatario_numero")
    locatario_bairro = request.form.get("locatario_bairro")
    locatario_cep = request.form.get("locatario_cep")
    locatario_cidade = request.form.get("locatario_cidade")
    locatario_uf = request.form.get("locatario_uf")

    carro_id = request.form.get("carro")
    data_inicio_str = request.form.get("data_inicio")
    data_fim_str = request.form.get("data_fim")

    data_inicio = datetime.strptime(data_inicio_str, "%Y-%m-%d").date()
    data_fim = datetime.strptime(data_fim_str, "%Y-%m-%d").date()

    if data_fim <= data_inicio:
        dias_locacao = 1
        data_fim = data_inicio + timedelta(days=1)
    else:
        dias_locacao = (data_fim - data_inicio).days

    # Arquivos anexos
    arquivo_cnh = request.files.get("arquivo_cnh")
    arquivo_comprovante = request.files.get("arquivo_comprovante")

    caminho_cnh = salvar_upload(arquivo_cnh)
    caminho_comprovante = salvar_upload(arquivo_comprovante)

    carro = CARROS[carro_id]

    # ---------- 2. PREPARA CONTEXTO PARA O DOCX ----------
    contexto = {
        "locatario_nome": up(locatario_nome),
        "locatario_nacionalidade": up(locatario_nacionalidade),
        "locatario_estado_civil": up(locatario_estado_civil),
        "locatario_profissao": up(locatario_profissao),
        "locatario_rg": locatario_rg,
        "locatario_cpf": locatario_cpf,
        "locatario_cnh": locatario_cnh,

        "locatario_rua": up(locatario_rua),
        "locatario_numero": locatario_numero,
        "locatario_bairro": up(locatario_bairro),
        "locatario_cep": locatario_cep,
        "locatario_cidade": up(locatario_cidade),
        "locatario_uf": up(locatario_uf),

        "carro_marca": up(carro["marca"]),
        "carro_modelo": up(carro["modelo"]),
        "carro_ano": carro["ano"],
        "carro_cor": up(carro["cor"]),
        "carro_placa": up(carro["placa"]),
        "carro_categoria": up(carro["categoria"]),
        "carro_valor_avaliacao": carro["valor_avaliacao"],

        "dias_locacao": dias_locacao,
        "data_inicio": data_inicio.strftime("%d/%m/%Y"),
        "data_fim": data_fim.strftime("%d/%m/%Y"),
    }

    # ---------- 3. GERA DOCX A PARTIR DO TEMPLATE ----------
    doc = DocxTemplate(TEMPLATE_DOCX)
    doc.render(contexto)

    nome_base = f"contrato_{locatario_nome.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d%H%M%S')}"
    caminho_docx_preenchido = os.path.join(OUTPUT_FOLDER, f"{nome_base}.docx")

    doc.save(caminho_docx_preenchido)

    # ---------- 4. ENVIA CONTRATO (DOCX) + ANEXOS POR E-MAIL ----------
    assunto = f"Pré-contrato de locação - {locatario_nome}"
    corpo_email = f"""
Foi gerada uma pré-solicitação de contrato de locação.

Locatário: {locatario_nome}
CPF: {locatario_cpf}
Carro: {carro['nome_exibicao']}
Período: {contexto['data_inicio']} até {contexto['data_fim']}
Quantidade de dias: {dias_locacao}

Anexos:
- Contrato em DOCX para conferência e assinatura.
- CNH do cliente.
- Comprovante de endereço do cliente.
"""

    anexos = [caminho_docx_preenchido]
    if caminho_cnh:
        anexos.append(caminho_cnh)
    if caminho_comprovante:
        anexos.append(caminho_comprovante)

    status_envio = "ok"
    erro_envio = None

    try:
        enviar_email_com_anexos(
            assunto=assunto,
            corpo=corpo_email,
            destinatario=EMAIL_DESTINO,
            caminhos=anexos,
        )
    except Exception as e:
        status_envio = "erro"
        erro_envio = str(e)

    # ---------- 5. MOSTRA TELA DE CONFIRMAÇÃO ----------
    return render_template(
        "sucesso.html",
        status_envio=status_envio,
        erro_envio=erro_envio,
        locatario_nome=locatario_nome,
        carro_nome=carro["nome_exibicao"],
        data_inicio=contexto["data_inicio"],
        data_fim=contexto["data_fim"],
        dias_locacao=dias_locacao,
        email_destino=EMAIL_DESTINO,
    )


if __name__ == "__main__":
    app.run(debug=True)
