import os
import time
import urllib.parse
from datetime import datetime

import pandas as pd
from flask import Flask, request, render_template_string
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from werkzeug.utils import secure_filename
from selenium.webdriver.common.action_chains import ActionChains


# ----------------- CONFIG BÁSICA -----------------

NOME_PLANILHA = "Planilha1"
COLUNA_NOME = "Nome do Cliente"
COLUNA_TEL = "Número do Celular"

app = Flask(__name__)

BASE_DIR = os.path.dirname(__file__)
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# ----------------- TEMPLATE HTML -----------------

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Disparador Automático</title>

    <style>
        body {
            margin: 0;
            padding: 0;
            font-family: "Segoe UI", Arial, sans-serif;
            background: #0c0c0c;
            color: #e8e8e8;
        }

        .background {
            width: 100%;
            min-height: 100vh;
            background-image: url('{{ url_for("static", filename="bg.png") }}');
            background-size: cover;
            background-position: center;
            background-repeat: no-repeat;
            padding: 18px 15px;
            display: flex;
            justify-content: center;
        }

        .container {
            max-width: 1000px;
            width: 95%;
        }

        .logo-top {
            text-align: center;
            margin-bottom: 10px;
        }

        .app-title {
            margin: 0;
            font-size: 26px;
            font-weight: 700;
            color: #e8ffe8;
            text-shadow: 0 0 6px rgba(0,255,140,0.6);
        }

        .app-subtitle {
            margin: 4px 0 0;
            font-size: 12px;
            color: #9bb;
        }

        .card {
            background: rgba(20, 20, 20, 0.75);
            padding: 22px 24px;
            margin-bottom: 16px;
            border-radius: 18px;
            box-shadow: 0 0 22px rgba(0, 255, 120, 0.18);
            backdrop-filter: blur(10px);
            border: 1px solid rgba(0, 255, 120, 0.16);
        }

        .card h2 {
            margin: 0;
            font-size: 19px;
            color: #00ff7f;
            text-shadow: 0 0 6px rgba(0,255,120,0.35);
        }

        .card-header {
            display: flex;
            align-items: center;
            gap: 12px;
            margin-bottom: 14px;
        }

        .card-accent {
            width: 4px;
            height: 40px;
            border-radius: 999px;
            background: linear-gradient(180deg, #00ff7f, #00b35a);
            box-shadow: 0 0 12px rgba(0, 255, 120, 0.7);
        }

        .card-header-text {
            display: flex;
            flex-direction: column;
            gap: 2px;
        }

        .card-subtitle {
            margin: 0;
            font-size: 12px;
            color: #9aa;
        }

        label {
            display: block;
            margin-top: 18px;
            margin-bottom: 6px;
            font-weight: 600;
            color: #cfcfcf;
            font-size: 14px;
        }

        input[type="text"],
        input[type="number"],
        textarea {
            width: 100%;
            padding: 14px;
            background: rgba(255, 255, 255, 0.06);
            border: 1px solid rgba(0, 255, 100, 0.35);
            border-radius: 12px;
            color: #e8e8e8;
            font-size: 15px;
            outline: none;
            transition: 0.2s;
        }

        input:focus,
        textarea:focus {
            border-color: #00ff7f;
            box-shadow: 0 0 12px rgba(0, 255, 120, 0.45);
            background: rgba(255, 255, 255, 0.11);
        }

        textarea {
            height: 150px;
            resize: vertical;
        }

        .info {
            font-size: 12px;
            color: #888;
            margin-top: 3px;
            margin-bottom: 8px;
        }

        /* Estilização do botão de upload de arquivo */
        input[type="file"] {
            width: 100%;
            padding: 14px;
            border-radius: 12px;
            cursor: pointer;
            background: rgba(255, 255, 255, 0.05);
            border: 1px solid rgba(0, 255, 100, 0.35);
            color: #ccc;
            font-size: 15px;
            transition: 0.25s ease;
        }

        input[type="file"]::-webkit-file-upload-button {
            visibility: hidden;
        }

        input[type="file"]::before {
            content: "Selecionar arquivo";
            display: inline-block;
            background: rgba(0, 255, 100, 0.2);
            border: 1px solid rgba(0, 255, 100, 0.5);
            padding: 10px 18px;
            margin-right: 12px;
            font-size: 15px;
            font-weight: 600;
            border-radius: 10px;
            color: #00ff9d;
            cursor: pointer;
            transition: 0.25s ease;
        }

        input[type="file"]:hover::before {
            background: rgba(0, 255, 100, 0.35);
            box-shadow: 0 0 12px rgba(0, 255, 100, 0.5);
        }

        input[type="file"]:active::before {
            background: rgba(0, 255, 100, 0.55);
        }

        button {
            width: 100%;
            padding: 16px;
            background: #00ff7f;
            border: none;
            color: #000;
            font-size: 18px;
            font-weight: bold;
            border-radius: 14px;
            cursor: pointer;
            margin-top: 10px;
            margin-bottom: 6px;
            box-shadow: 0 0 15px rgba(0, 255, 120, 0.5);
            transition: 0.25s ease;
        }

        button:hover {
            background: #00ffaa;
            box-shadow: 0 0 25px rgba(0, 255, 140, 0.8);
        }

        .log-box {
            background: rgba(0, 0, 0, 0.35);
            border: 1px solid rgba(0,255,100,0.25);
            border-radius: 14px;
            padding: 15px;
            white-space: pre-wrap;
            font-family: Consolas, monospace;
            max-height: 350px;
            overflow-y: auto;
            color: #d4ffd4;
            font-size: 13px;
        }

        .cards-grid {
            display: grid;
            grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
            gap: 20px;
            margin-bottom: 16px;
        }

        /* ===== RESPONSIVIDADE ===== */

        @media (max-width: 900px) {
            .background {
                padding: 12px 8px;
            }

            .container {
                max-width: 100%;
                width: 100%;
            }

            .card {
                padding: 16px 14px;
                margin-bottom: 12px;
                border-radius: 14px;
            }

            .app-title {
                font-size: 22px;
            }

            button {
                padding: 14px;
                font-size: 16px;
                border-radius: 12px;
            }
        }

        @media (max-width: 800px) {
            .cards-grid {
                grid-template-columns: 1fr;
                gap: 10px;
                margin-bottom: 12px;
            }
        }

        @media (max-width: 480px) {
            body {
                font-size: 14px;
            }

            .background {
                min-height: unset;
                padding: 8px 4px;
                padding-bottom: 20px;
            }

            .logo-top {
                margin-bottom: 6px;
            }

            .app-title {
                font-size: 20px;
            }

            .app-subtitle {
                font-size: 11px;
            }

            .card {
                padding: 12px 10px;
                border-radius: 12px;
                box-shadow: 0 0 14px rgba(0, 255, 120, 0.2);
            }

            label {
                font-size: 13px;
            }

            input[type="text"],
            input[type="number"],
            textarea,
            input[type="file"] {
                font-size: 14px;
                padding: 10px;
            }

            input[type="file"]::before {
                font-size: 14px;
                padding: 8px 14px;
                margin-right: 8px;
            }

            button {
                margin-top: 8px;
                margin-bottom: 0px;
                padding: 12px;
                font-size: 15px;
            }

            .log-box {
                max-height: 240px;
                font-size: 12px;
            }
        }

        /* Garantir que os inputs não “estourem” o card */
        input[type="text"],
        input[type="number"],
        input[type="file"],
        textarea {
            box-sizing: border-box;
            max-width: 100%;
        }

        /* Se ainda assim sobrar 1px, o card corta */
        .card {
            overflow: hidden;
        }
    </style>
</head>

<body>
<div class="background">
<div class="container">

    <div class="logo-top">
        <h1 class="app-title">Disparador Automático</h1>
        <p class="app-subtitle">Envio inteligente de mensagens via WhatsApp Web</p>
    </div>

    <form method="post" enctype="multipart/form-data">

        <!-- LINHA 1: PLANILHA + CONFIG EM GRID -->
        <div class="cards-grid">

            <!-- CARTÃO 1: PLANILHA -->
            <div class="card">
                <div class="card-header">
                    <div class="card-accent"></div>
                    <div class="card-header-text">
                        <h2>Planilha</h2>
                        <p class="card-subtitle">Selecione o arquivo com a base de clientes.</p>
                    </div>
                </div>

                <label>Selecione a planilha (.xlsx):</label>
                <input type="file" name="arquivo_excel_upload" accept=".xlsx" required>

                <div class="info">
                    Arquivo selecionado: {{ arquivo_excel or 'nenhum arquivo enviado ainda' }}<br>
                    A aba deve ser '{{ nome_planilha }}' e conter as colunas '{{ coluna_nome }}' e '{{ coluna_tel }}'.
                </div>
            </div>

            <!-- CARTÃO 2: CONFIGURAÇÕES -->
            <div class="card">
                <div class="card-header">
                    <div class="card-accent"></div>
                    <div class="card-header-text">
                        <h2>Configurações</h2>
                        <p class="card-subtitle">Defina o intervalo entre um envio e outro.</p>
                    </div>
                </div>

                <label>Intervalo entre envios (segundos):</label>
                <input type="number" min="0.5" step="0.1" name="intervalo" value="{{ intervalo or '1.0' }}">
            </div>

        </div> <!-- fim cards-grid -->

        <!-- CARTÃO 3: MENSAGEM -->
        <div class="card">
            <div class="card-header">
                <div class="card-accent"></div>
                <div class="card-header-text">
                    <h2>Mensagem</h2>
                    <p class="card-subtitle">Conteúdo que será enviado para cada contato.</p>
                </div>
            </div>

            <label>Conteúdo da mensagem (use {NOME}):</label>
            <textarea name="mensagem">{{ mensagem or "Olá {NOME}, tudo bem?\\nGostaria de conversar sobre a organização da sua vida financeira." }}</textarea>

            <label style="margin-top: 16px;">Anexo (opcional – imagem ou PDF):</label>
            <input type="file" name="arquivo_anexo" accept=".png,.jpg,.jpeg,.pdf">

            <div class="info">
                Se você anexar um arquivo, ele será enviado junto com a mensagem
                para todos os clientes desta base.
            </div>
        </div>

        <!-- BOTÃO DE DISPARO -->
        <button type="submit">Iniciar Disparos</button>

        {% if logs %}
        <div class="card">
            <div class="card-header">
                <div class="card-accent"></div>
                <div class="card-header-text">
                    <h2>Log do processo</h2>
                    <p class="card-subtitle">Acompanhe o resultado dos envios.</p>
                </div>
            </div>
            <div class="log-box">{{ logs }}</div>
        </div>
        {% endif %}

    </form>

</div>
</div>
</body>
</html>
"""

# ----------------- FUNÇÃO DE LOG -----------------


def log(msg, logs_list):
    ts = datetime.now().strftime("%H:%M:%S")
    linha = f"[{ts}] {msg}"
    print(linha)
    logs_list.append(linha)


# ----------------- LÓGICA DE ENVIO -----------------
from selenium.webdriver.common.action_chains import ActionChains

def processar_envios(arquivo, mensagem_digitada, intervalo, logs_list, caminho_anexo=None):
    if not os.path.exists(arquivo):
        log(f"ERRO: Arquivo não encontrado: {arquivo}", logs_list)
        return

    # Leitura do Excel
    try:
        df = pd.read_excel(arquivo, sheet_name=NOME_PLANILHA)
        df.columns = df.columns.str.strip()

        if COLUNA_NOME not in df.columns or COLUNA_TEL not in df.columns:
            log(f"ERRO: Faltam colunas '{COLUNA_NOME}' ou '{COLUNA_TEL}'", logs_list)
            return

        if "Status" not in df.columns:
            df["Status"] = ""
        df["Status"] = df["Status"].astype(str)

    except Exception as e:
        log(f"ERRO Excel: {e}", logs_list)
        return

    # Conta quantos serão enviados
    validos = 0
    for _, row in df.iterrows():
        raw_telefone = row.get(COLUNA_TEL, "")
        status_row = str(row.get("Status", "")).strip()

        try:
            if pd.isna(raw_telefone):
                telefone_tmp = ""
            else:
                telefone_tmp = str(int(float(raw_telefone)))
        except Exception:
            telefone_tmp = str(raw_telefone).strip()

        if status_row != "Enviado" and telefone_tmp not in ["nan", "", "0"]:
            validos += 1

    log(f"Processando {validos} clientes válidos...", logs_list)

    # Abre Chrome + WhatsApp Web
    try:
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=options
        )
        driver.maximize_window()
        driver.get("https://web.whatsapp.com")
        log("WhatsApp Web aberto. Faça login se necessário...", logs_list)

        try:
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'div[contenteditable="true"]'))
            )
            log("WhatsApp Web carregado. Iniciando disparos...", logs_list)
        except Exception:
            log("AVISO: Não foi possível confirmar carregamento do WhatsApp Web.", logs_list)

    except Exception as e:
        log(f"ERRO Driver: {e}", logs_list)
        return

    enviados = 0
    erros = 0

    # ========================= LOOP PRINCIPAL =========================
    for index, row in df.iterrows():
        nome = str(row.get(COLUNA_NOME, "")).strip()
        raw_telefone = row.get(COLUNA_TEL, "")
        status = str(row.get("Status", "")).strip()

        # Trata telefone
        try:
            if pd.isna(raw_telefone):
                telefone = ""
            else:
                telefone = str(int(float(raw_telefone)))
        except Exception:
            telefone = str(raw_telefone).strip()

        if status == "Enviado" or telefone in ["nan", "", "0"]:
            continue

        log(f"Enviando para: {nome} ({telefone})...", logs_list)

        try:
            mensagem_final = mensagem_digitada.replace("{NOME}", nome)
            texto_codificado = urllib.parse.quote(mensagem_final)

            if caminho_anexo:
                link = f"https://web.whatsapp.com/send?phone={telefone}"
            else:
                link = f"https://web.whatsapp.com/send?phone={telefone}&text={texto_codificado}"

            driver.get(link)
            wait = WebDriverWait(driver, 30)

            # Passo 1 — espera caixa de digitação
            log("   -> Passo 1: esperando caixa de mensagem...", logs_list)
            try:
                caixa_msg = wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "div[contenteditable='true'][data-tab='10']")
                    )
                )
            except Exception:
                caixa_msg = wait.until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, "div[contenteditable='true']")
                    )
                )

            # ========================= COM ANEXO =========================
            if caminho_anexo:
                try:
                    # Passo 2 — digitar mensagem
                    log("   -> Passo 2: digitando mensagem no chat...", logs_list)
                    try:
                        caixa_msg.click()
                    except Exception:
                        pass
                    caixa_msg.send_keys(mensagem_final)

                    # Passo 3 — clicar no botão +
                    log("   -> Passo 3: clicando no botão de anexar por coordenada...", logs_list)
                    actions = ActionChains(driver)
                    try:
                        actions.move_to_element_with_offset(caixa_msg, -70, 0).click().perform()
                    except Exception:
                        actions.move_to_element_with_offset(caixa_msg, -90, 0).click().perform()
                    time.sleep(1.0)

                    # Passo 4 — localizar input genérico
                    log("   -> Passo 4: procurando input[type='file']...", logs_list)
                    file_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='file']")
                    log(f"   -> Passo 4: {len(file_inputs)} input[type='file'] encontrados.", logs_list)

                    if not file_inputs:
                        raise Exception("Nenhum input[type='file'] encontrado.")

                    file_input = file_inputs[-1]

                    log(f"   -> Passo 5: enviando caminho do arquivo: {caminho_anexo}", logs_list)
                    file_input.send_keys(caminho_anexo)

                    # Passo 6 — esperar preview
                    log("   -> Passo 6: aguardando preview do anexo...", logs_list)
                    time.sleep(3.0)

                    # Passo 7 — clicar no botão verde OU ENTER (fallback)
                    log("   -> Passo 7: procurando botão de enviar...", logs_list)
                    send_clicked = False

                    try:
                        # tenta vários seletores de botão enviar
                        send_btn = WebDriverWait(driver, 12).until(
                            EC.element_to_be_clickable(
                                (
                                    By.CSS_SELECTOR,
                                    "button[data-testid='send'], "
                                    "div[role='button'][aria-label*='Enviar'], "
                                    "button[aria-label*='Enviar']"
                                )
                            )
                        )
                        driver.execute_script("arguments[0].click();", send_btn)
                        send_clicked = True
                        log("   -> Passo 7: botão de enviar clicado (seletores padrão).", logs_list)
                    except Exception as e1:
                        log(f"   -> Passo 7: botão padrão não encontrado: {repr(e1)}", logs_list)
                        try:
                            # tenta achar o ícone 'send' e subir pro botão
                            span_icon = WebDriverWait(driver, 5).until(
                                EC.presence_of_element_located(
                                    (By.CSS_SELECTOR, "span[data-icon='send']")
                                )
                            )
                            try:
                                send_btn2 = span_icon.find_element(
                                    By.XPATH, "./ancestor::div[@role='button'][1]"
                                )
                            except Exception:
                                send_btn2 = span_icon
                            driver.execute_script("arguments[0].click();", send_btn2)
                            send_clicked = True
                            log("   -> Passo 7: botão de enviar clicado via ícone 'send'.", logs_list)
                        except Exception as e2:
                            log(f"   -> Passo 7: não achei ícone 'send': {repr(e2)}", logs_list)
                            send_clicked = False

                    if not send_clicked:
                        # Fallback: ENTER na legenda/caixa ativa
                        log("   -> Passo 7 (fallback): tentando ENTER na legenda...", logs_list)
                        try:
                            try:
                                legenda = driver.find_element(
                                    By.CSS_SELECTOR,
                                    "div[contenteditable='true'][data-tab='10']"
                                )
                            except Exception:
                                legenda = caixa_msg
                            legenda.send_keys(Keys.ENTER)
                            log("   -> ENTER enviado na legenda (fallback).", logs_list)
                        except Exception as e3:
                            raise Exception(
                                f"Falha ao enviar (nem botão nem ENTER funcionaram): {e3}"
                            )

                except Exception as e:
                    log(f"   -> ERRO no fluxo de anexo: {repr(e)}", logs_list)
                    df.at[index, "Status"] = "Erro"
                    erros += 1
                    log(f"   -> Marcado como ERRO. (Enviados: {enviados} | Erros: {erros})", logs_list)
                    continue

                # OK
                time.sleep(max(intervalo, 0.5))
                df.at[index, "Status"] = "Enviado"
                enviados += 1
                log(f"   -> Sucesso! (Enviados: {enviados} | Erros: {erros})", logs_list)
                continue

            # ========================= SEM ANEXO =========================
            else:
                time.sleep(0.5)
                try:
                    caixa_msg.click()
                except Exception:
                    pass

                driver.switch_to.active_element.send_keys(Keys.ENTER)

                time.sleep(max(intervalo, 0.5))

                df.at[index, "Status"] = "Enviado"
                enviados += 1
                log(f"   -> Sucesso! (Enviados: {enviados} | Erros: {erros})", logs_list)

        except Exception as e:
            log(f"   -> Falha ao enviar: {repr(e)}", logs_list)
            df.at[index, "Status"] = "Erro"
            erros += 1
            log(f"   -> Marcado como ERRO. (Enviados: {enviados} | Erros: {erros})", logs_list)

    # salva planilha
    try:
        df.to_excel(arquivo, sheet_name=NOME_PLANILHA, index=False)
        log("Planilha atualizada com o status dos envios.", logs_list)
    except Exception as e:
        log(f"ERRO ao salvar a planilha: {e}", logs_list)

    driver.quit()
    log("Processo finalizado.", logs_list)

# ----------------- ROTAS FLASK -----------------

@app.route("/", methods=["GET", "POST"])
def index():
    logs_list = []
    arquivo_excel = ""
    intervalo = "1.0"
    mensagem = ""
    caminho_anexo = None

    if request.method == "POST":
        file = request.files.get("arquivo_excel_upload")
        arquivo_anexo = request.files.get("arquivo_anexo")
        intervalo_str = request.form.get("intervalo", "1.0").strip()
        mensagem = request.form.get("mensagem", "").strip()

        if file and file.filename:
            filename = secure_filename(file.filename)
            save_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(save_path)
            arquivo_excel = filename

            # anexo opcional
            if arquivo_anexo and arquivo_anexo.filename:
                anexoname = secure_filename(arquivo_anexo.filename)
                anexo_path = os.path.join(app.config["UPLOAD_FOLDER"], anexoname)
                arquivo_anexo.save(anexo_path)
                caminho_anexo = os.path.abspath(anexo_path)
            else:
                caminho_anexo = None

            try:
                intervalo_val = float(intervalo_str)
                if intervalo_val <= 0:
                    intervalo_val = 1.0
            except Exception:
                intervalo_val = 1.0

            processar_envios(save_path, mensagem, intervalo_val, logs_list, caminho_anexo)
            intervalo = str(intervalo_val)
        else:
            logs_list.append("Nenhum arquivo foi enviado.")

    logs_text = "\n".join(logs_list) if logs_list else ""

    return render_template_string(
        HTML_TEMPLATE,
        logs=logs_text,
        arquivo_excel=arquivo_excel,
        intervalo=intervalo,
        mensagem=mensagem,
        nome_planilha=NOME_PLANILHA,
        coluna_nome=COLUNA_NOME,
        coluna_tel=COLUNA_TEL,
    )


# ----------------- MAIN (PARA RODAR COMO PROGRAMA / EXE) -----------------


if __name__ == "__main__":
    import threading
    import webbrowser

    def abrir_navegador():
        time.sleep(1)
        webbrowser.open("http://127.0.0.1:5000")

    threading.Thread(target=abrir_navegador, daemon=True).start()
    app.run(host="127.0.0.1", port=5000, debug=False)
