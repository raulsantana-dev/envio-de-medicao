import asyncio
import pandas as pd
import os
import shutil
from datetime import datetime
from utils.excel_hander import gerar_planilha_para_cliente, extrair_links_reais
from utils.envio_email import enviar_email
from dotenv import load_dotenv
from database.snowflake_conector import insere_cliente

load_dotenv()

async def processar_arquivo(caminho_andamento, caminho_finalizado, nome_arquivo, caminho_origem):
    try:
        mes_ano = datetime.now().strftime('%B/%Y').capitalize()
        df = extrair_links_reais(caminho_andamento)
        clientes = df['cliente'].unique()

        for cliente_nome in clientes:
            print(f"🔧 Processando cliente: {cliente_nome}")
            df_cliente = df[df['cliente'] == cliente_nome]

            # Coleta todos os links únicos e válidos para o cliente
            links = df_cliente['Link'].dropna().unique().tolist()
            texto_links = "\n".join(links) if links else "Nenhum link encontrado."

            # Validação
            validacao = df_cliente['Validação'].dropna().unique()[0] if not df_cliente['Validação'].dropna().empty else ""

            # Captura todos os e-mails da célula (pode estar separado por , ou ;)
            email_raw = df_cliente['Email'].dropna().values[0] if not df_cliente['Email'].dropna().empty else None
            if email_raw:
                destinatario = [e.strip() for e in email_raw.replace(';', ',').split(',') if e.strip()]
            else:
                print(f"⚠️ Nenhum e-mail encontrado para {cliente_nome}, pulando envio.")
                continue

            # Envia dados ao banco
            for _, linha in df_cliente.iterrows():
                cliente = linha['cliente'] if pd.notna(linha['cliente']) else ''
                cnpj = linha['CNPJ/CPF'] if pd.notna(linha['CNPJ/CPF']) else ''
                chassi = linha['Chassi'] if pd.notna(linha['Chassi']) else ''
                placa = linha['Placa 1'] if pd.notna(linha['Placa 1']) else ''
                codigo_infracao = linha['Cód. Da Infração'] if pd.notna(linha['Cód. Da Infração']) else ''
                ait = linha['AIT'] if pd.notna(linha['AIT']) else ''
                print(f"{cliente} | {cnpj} | {chassi} | {placa} | {codigo_infracao} | {ait} | {mes_ano}")
                insere_cliente(cliente, cnpj, chassi, ait, mes_ano, placa)

            # Gera planilha para o cliente
            caminho_arquivo = gerar_planilha_para_cliente(cliente_nome, df_cliente.to_dict(orient="records"))
            print(f"📝 Planilha gerada para {cliente_nome}: {caminho_arquivo}")

            # Envia e-mail
            await asyncio.sleep(1)
            await enviar_email(destinatario, cliente_nome, caminho_arquivo, texto_links, validacao)
            print(f"📧 E-mail enviado para {destinatario}")

            # Remove a planilha temporária
            try:
                if os.path.isfile(caminho_arquivo):
                    os.remove(caminho_arquivo)
                    print(f"🗑️ Planilha removida: {caminho_arquivo}")
            except Exception as e:
                print(f"❌ Erro ao remover planilha {caminho_arquivo}: {e}")

        print(f"✅ Arquivo {nome_arquivo} finalizado com sucesso.")

    except Exception as e:
        print(f"❌ Erro ao processar cria_planilha {nome_arquivo}: {e}")
        shutil.move(caminho_andamento, caminho_origem)

    else:
        # Move o arquivo para a pasta Finalizados com o mesmo nome original
      data_hoje = datetime.now()
    # Monta nomes das pastas
    mes = data_hoje.strftime("(%m) - %B_%Y").lower()  # exemplo: (06) - junho_2025
    dia = data_hoje.strftime("%d_%m_%Y")              # exemplo: 18_06_2025

    # Cria o caminho final com subpastas
    caminho_finalizado_final = os.path.join(caminho_finalizado, mes, dia)
    nome_arquivo = os.path.basename(caminho_andamento)
    # Garante que os diretórios existem
    os.makedirs(caminho_finalizado_final, exist_ok=True)
    # Caminho completo de destino
    caminho_destino = os.path.join(caminho_finalizado_final, nome_arquivo)

    # Move o arquivo
    shutil.move(caminho_andamento, caminho_destino)

    print(f"📦 Arquivo movido para '{caminho_finalizado_final}': {nome_arquivo}")


# adição de trecho antigo para aprendizado, desconsiderar caso for executar o codigo #




# import os
# import ssl
# import base64
# import getpass
# import httpx
# from dotenv import load_dotenv

# load_dotenv()

# CERT_FILE = "C:/Projects/Vamos_EnvioMedicao/certificados/VAMOSLOCACAO_23373000000132.crt"
# KEY_FILE = "C:/Projects/Vamos_EnvioMedicao/key/VAMOSLOCACAO_23373000000132.key"
# PASS_PHRASE = os.getenv("senhaKey")

# pasta_destino = "C:/Projects/Vamos_EnvioMedicao/multas"  # atenção à barra invertida, troque por /

# async def download_multa(placa, codigoOrgao, ait, codigoInfracao, url_api):
#     url_get = f'{url_api}/consultas/sne/pdf/placa/{placa}/codigoOrgao/{codigoOrgao}/numeroAit/{ait}/codigoInfracao/{codigoInfracao}/NA'

#     # Criar contexto SSL com passphrase
#     ssl_context = ssl.create_default_context(ssl.Purpose.SERVER_AUTH)
#     ssl_context.load_cert_chain(certfile=CERT_FILE, keyfile=KEY_FILE, password=PASS_PHRASE)
#     async with httpx.AsyncClient(verify=ssl_context) as client:
#         response = await client.get(url_get)

#     if response.status_code == 200:
#         retorno_api = response.json()
#         base64_pdf = retorno_api.get("base64")

#         if base64_pdf:
#             pdf_bytes = base64.b64decode(base64_pdf)

#             os.makedirs(pasta_destino, exist_ok=True)

#             caminho_pdf = os.path.join(pasta_destino, f"{placa}_{codigoInfracao}_{ait}.pdf")

#             with open(caminho_pdf, "wb") as f:
#                 f.write(pdf_bytes)

#             print("✅ PDF salvo com sucesso:", caminho_pdf)
#         else:
#             print("⚠️ Nenhum conteúdo base64 retornado.")
#     else:
#         print("❌ Erro na requisição:", response.status_code, response.text)
