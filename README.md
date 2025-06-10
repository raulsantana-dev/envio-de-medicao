# ENVIO MEDICAO v1.0.0

Projeto em **Python 3** para automação do processo de envio de medições. Utiliza integração com **Snowflake**, manipulação de **Excel**, automação de envio de e-mails e processamento de arquivos com evidências.

---

## ✅ Tecnologias utilizadas

- **Python 3.10+**
- **Snowflake Connector**
- **openpyxl / pandas**
- **smtplib**
- **dotenv**
- **os / shutil / glob**

---

## ✅ Estrutura do Projeto
O projeto é estruturado de forma modular, com divisão clara em camadas para facilitar manutenção, escalabilidade e legibilidade.


# ✅ Funcionalidades

- ✅ Conexão com banco de dados **Snowflake**
- ✅ Leitura e estruturação de arquivos **Excel**
- ✅ Processamento e renomeação de arquivos de evidência
- ✅ Envio automático de e-mail com os documentos anexados
- ✅ Modularização para reuso e fácil manutenção

---

## ✅ Instalação

```bash
git clone <seu-repositorio>
cd Vamos_EnvioMedicao
python -m venv venv
source venv/bin/activate  # Ou venv\Scripts\activate no Windows
pip install -r requirements.txt
