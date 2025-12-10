# 📤 Disparador de Mensagens no WhatsApp com Anexo  
Automação completa de envio de mensagens personalizadas via WhatsApp Web utilizando **Python**, **Flask**, **Selenium** e **webdriver-manager**.  
Este projeto permite carregar uma planilha, escrever mensagens personalizadas e enviar textos e anexos automaticamente para uma lista de contatos.

---

## 🚀 Funcionalidades

- Interface Web simples e intuitiva (Flask)
- Envio automatizado de mensagens via WhatsApp Web (Selenium)
- Suporte a anexo de imagens ou documentos
- Leitura de planilha Excel com **pandas**
- Processamento individual dos contatos
- Feedback em tempo real no console (log de envio)
- Contagem de enviados / erros
- Upload dinâmico de arquivos
- Ambiente isolado com `venv`

---

## 🛠 Tecnologias Utilizadas

- **Python 3.10+**
- **Flask**
- **Selenium WebDriver**
- **webdriver-manager**
- **pandas**
- **openpyxl**
- HTML + CSS (interface)
- Chrome + ChromeDriver

---

## 📁 Estrutura do Projeto

Disparador_Whatsapp_Anexo/
│── app_web.py # Aplicação Flask + lógica de automação
│── .gitignore
│── static/
│ ├── bg.png # Imagem de fundo
│ └── whatsapp_neon.png # Ícone ou imagem usada na interface
│── uploads/ # Arquivos e planilhas enviados pelo usuário
│── build/ # (gerado ao compilar para .exe)
│── dist/ # (gerado ao compilar para .exe)
│── venv/ # Ambiente virtual Python

yaml
Copiar código

---

## 🧩 Instalação e Configuração

### 1️⃣ Clone o repositório

```bash
git clone https://github.com/ThiagoNatario/Disparador_Whatsapp_Anexo.git
cd Disparador_Whatsapp_Anexo
2️⃣ Criar ambiente virtual
bash
Copiar código
python -m venv venv
Ativar:

Windows Powershell

bash
Copiar código
.\venv\Scripts\activate
3️⃣ Instalar dependências
bash
Copiar código
pip install flask selenium webdriver-manager pandas openpyxl Werkzeug
4️⃣ Instalar o Google Chrome
Baixe aqui (site oficial):
https://www.google.com/chrome/

5️⃣ Executar a aplicação
bash
Copiar código
python app_web.py
Acesse no navegador:

cpp
Copiar código
http://127.0.0.1:5000
📄 Formato da Planilha Esperada
A planilha deve conter:

Uma coluna com o nome do cliente

Uma coluna com o número de telefone no formato internacional 55DDDNUMERO

Exemplo:

Nome do Cliente	Número do Celular
João Silva	5521999999999
Maria Souza	5531988888888

🧪 Funcionamento do Envio
O usuário carrega a planilha no navegador

A interface exibe os contatos detectados

O sistema abre o WhatsApp Web automaticamente

É necessário escanear o QR Code uma única vez

O Selenium envia a mensagem individualmente para cada número

Se houver anexo, é enviado antes da mensagem

Ao final, o console mostra um resumo:

makefile
Copiar código
Enviados: 10
Erros: 0
Processo finalizado.
⚠️ Limitações e Avisos
Este projeto é para fins educacionais

Automação no WhatsApp pode resultar em limite temporário se utilizada com volume excessivo

Não utilize para spam

📌 Próximos Melhoramentos (roadmap)
Interface mais moderna (HTML/CSS revisado)

Upload de múltiplos anexos

Interface com preview da mensagem

Logs em arquivo .txt

Dashboard com histórico dos envios

Versão .exe portable

👨‍💻 Autor
Thiago Natario
Repositório GitHub: https://github.com/ThiagoNatario
Projeto desenvolvido com auxílio do ChatGPT.

⭐ Contribuições
Sinta-se à vontade para abrir issues e pull requests.
Se o projeto for útil, deixe uma ⭐ no GitHub!
