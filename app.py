from flask import Flask, render_template, request
from datetime import datetime, timedelta
import os
import subprocess  # para chamar o LibreOffice

from werkzeug.utils import secure_filename
from docxtpl import DocxTemplate
from PyPDF2 import PdfMerger
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader

import smtplib
from email.message import EmailMessage

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


def imagem_para_pdf(caminho_imagem, caminho_pdf_saida):
    """Converte uma imagem (JPG, PNG, etc.) em um PDF de uma página."""
    c = canvas.Canvas(caminho_pdf_saida, pagesize=A4)
    largura_pagina, altura_pagina = A4

    img = ImageReader(caminho_imagem)
    largura_img, altura_img = img.getSize()

    escala = min(largura_pagina / largura_img, altura_pagina / altura_img)

    nova_largura = largura_img * escala
    nova_altura = altura_img * escala

    x = (largura_pagina - nova_largura) / 2
    y = (altura_pagina - nova_altura) / 2

    c.drawImage(img, x, y, nova_largura, nova_altura)
    c.showPage()
    c.save()


def docx_para_pdf(caminho_docx, caminho_pdf_destino):
    """
    Converte DOCX em PDF usando LibreOffice pela linha de comando (Linux).
    """
    pasta_saida = os.path.dirname(caminho_pdf_destino) or BASE_DIR

    resultado = subprocess.run(
        [
            "libreoffice",
            "--headless",
            "--convert-to", "pdf",
            "--outdir", pasta_saida,
            caminho_docx,
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )

    if resultado.returncode != 0:
        raise RuntimeError(f"Erro ao converter DOCX para PDF: {resultado.stderr}")

    # LibreOffice gera um PDF com o mesmo nome base do DOCX
    nome_base = os.path.splitext(os.path.basename(caminho_docx))[0] + ".pdf"
    pdf_gerado = os.path.join(pasta_saida, nome_base)

    # Se o destino for diferente do gerado, renomeia
    if pdf_gerado != caminho_pdf_destino:
        os.replace(pdf_gerado, caminho_pdf_destino)


def juntar_pdfs(lista_caminhos, caminho_saida):
    """
    Juntas os PDFs, ignorando arquivos inexistentes ou vazios (0 bytes)
    para evitar PyPDF2.errors.EmptyFileError.
    """
    merger = PdfMerger()

    for caminho in lista_caminhos:
        if not caminho:
            continue
        if not os.path.exists(caminho):
            continue
        if os.path.getsize(caminho) == 0:
            # se por algum motivo o PDF estiver vazio, pula
            continue

        merger.append(caminho)

    with open(caminho_saida, "wb") as f:
        merger.write(f)
    merger.close()


def enviar_email_com_anexo(assunto, corpo, destinatario, caminho_anexo):
    """Envia e-mail com o PDF de contrato anexado."""
    msg = EmailMessage()
    msg["Subject"] = assunto
    msg["From"] = SMTP_USER
    msg["To"] = destinatario
    msg.set_content(corpo)

    with open(caminho_anexo, "rb") as f:
        dados = f.read()
    nome_arquivo = os.path.basename(caminho_anexo)

    msg.add_attachment(
        dados,
        maintype="application",
        subtype="pdf",
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

    arquivo_cnh = request.files.get("arquivo_cnh")
    arquivo_comprovante = request.files.get("arquivo_comprovante")

    caminho_cnh = salvar_upload(arquivo_cnh)
    caminho_comprovante = salvar_upload(arquivo_comprovante)

    carro = CARROS[carro_id]

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
    caminho_pdf_contrato = os.path.join(OUTPUT_FOLDER, f"{nome_base}.pdf")

    doc.save(caminho_docx_preenchido)

    # ---------- 4. CONVERTE DOCX PARA PDF (LIBREOFFICE) ----------
    docx_para_pdf(caminho_docx_preenchido, caminho_pdf_contrato)

    pdfs_para_juntar = [caminho_pdf_contrato]

    # ---------- 5. TRATA ANEXOS (CNH + COMPROVANTE) ----------
    def preparar_pdf_anexo(caminho_arquivo):
        if not caminho_arquivo:
            return None
        ext = os.path.splitext(caminho_arquivo)[1].lower()
        if ext in [".jpg", ".jpeg", ".png"]:
            caminho_pdf_anexo = caminho_arquivo + ".pdf"
            imagem_para_pdf(caminho_arquivo, caminho_pdf_anexo)
            return caminho_pdf_anexo
        elif ext == ".pdf":
            return caminho_arquivo
        else:
            return None

    pdf_cnh = preparar_pdf_anexo(caminho_cnh)
    pdf_comp = preparar_pdf_anexo(caminho_comprovante)

    if pdf_cnh:
        pdfs_para_juntar.append(pdf_cnh)
    if pdf_comp:
        pdfs_para_juntar.append(pdf_comp)

    # ---------- 6. JUNTA TUDO EM UM ÚNICO PDF ----------
    caminho_pdf_final = os.path.join(OUTPUT_FOLDER, f"{nome_base}_FINAL.pdf")
    juntar_pdfs(pdfs_para_juntar, caminho_pdf_final)

    # ---------- 7. ENVIA O CONTRATO POR E-MAIL PARA A LOCADORA ----------
    assunto = f"Pré-contrato de locação - {locatario_nome}"
    corpo_email = f"""
Foi gerada uma pré-solicitação de contrato de locação.

Locatário: {locatario_nome}
CPF: {locatario_cpf}
Carro: {carro['nome_exibicao']}
Período: {contexto['data_inicio']} até {contexto['data_fim']}
Quantidade de dias: {dias_locacao}

O contrato completo segue em anexo para conferência e assinatura.
"""

    status_envio = "ok"
    erro_envio = None

    try:
        enviar_email_com_anexo(
            assunto=assunto,
            corpo=corpo_email,
            destinatario=EMAIL_DESTINO,
            caminho_anexo=caminho_pdf_final,
        )
    except Exception as e:
        status_envio = "erro"
        erro_envio = str(e)

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
