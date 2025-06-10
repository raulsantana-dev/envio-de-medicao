import os
import shutil
import pandas as pd
import asyncio
from dotenv import load_dotenv
from utils.processa_arquivos import processar_arquivo  # processar_arquivo deve ser async

load_dotenv()



# Caminhos das pastas
base_path = r"C:/Users/raul.araujo/JSL SA/Grupo Vamos - Execução Robôs - Medição - Envio de Medição"
pasta_executar = os.path.join(base_path, "Executar")
pasta_andamento = os.path.join(base_path, "Em andamento")
pasta_finalizados = os.path.join(base_path, "Finalizados")

# Garante que as pastas existem
os.makedirs(pasta_andamento, exist_ok=True)
os.makedirs(pasta_finalizados, exist_ok=True)

async def main():
    # Processa todos os arquivos Excel na pasta Executar
    for nome_arquivo in os.listdir(pasta_executar):
        if nome_arquivo.endswith(".xlsx"):
            print(f"📄 Iniciando processamento de: {nome_arquivo}")

            caminho_origem = os.path.join(pasta_executar, nome_arquivo)
            caminho_andamento = os.path.join(pasta_andamento, nome_arquivo)
            caminho_finalizado = os.path.join(pasta_finalizados, nome_arquivo)

            # Move para "Em andamento"
            shutil.move(caminho_origem, caminho_andamento)
            print("➡️  Arquivo movido para 'Em andamento'.")

            try:
                print("Processando os Arquivos")
                await processar_arquivo(caminho_andamento, caminho_finalizado, nome_arquivo, caminho_origem)
            except Exception as e:
                print(f"❌ Erro ao processar {nome_arquivo}: {e}")
                print(e)
                continue
    print("🏁 Todos os clientes foram processados.")

if __name__ == "__main__":
    asyncio.run(main())
