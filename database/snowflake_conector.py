import snowflake.connector
import os
from dotenv import load_dotenv

# Carrega as variáveis do .env
load_dotenv()

# Lê as credenciais do arquivo .env
loginSF = os.getenv("loginSF")
senhaSF = os.getenv("senhaSF")

def conectar_snowflake():
    """
    Estabelece conexão com o banco Snowflake.
    """
    conn = snowflake.connector.connect(
        account="sx27439.east-us-2.azure",
        user=loginSF,
        password=senhaSF,
        warehouse="PRODUCAO",
        role="RPA",
        database="_ENCLOSE",
        schema="RPA",
        client_session_keep_alive=True
    )
    return conn

def insere_cliente(cliente, cnpj_cpf, chassi, ait_principal, mes_ano, placa):
    """
    Executa uma query pré-definida com parâmetros dinâmicos.
    """
    query = """
       INSERT INTO RPA.MEDICOES_ENVIADAS (
           "Cliente", "CNPJ/CPF", "Chassi", "AitPrincipal", "Mes/Ano", "Placa"
       ) VALUES (%s, %s, %s, %s, %s, %s)
    """
    params = (cliente, cnpj_cpf, chassi, ait_principal, mes_ano, placa)

    conn = conectar_snowflake()
    cur = conn.cursor()
    try:
        cur.execute(query, params)
        conn.commit()
    finally:
        cur.close()
        conn.close()
