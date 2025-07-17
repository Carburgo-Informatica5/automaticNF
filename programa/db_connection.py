import os
import yaml
import oracledb
import logging

current_dir = os.path.dirname(os.path.abspath(__file__))

def load_db_config():
    with open(os.path.join(current_dir, "data", "oracar.yaml"), "r") as file:
        return yaml.safe_load(file)

def connect_to_db(db_config):
    if db_config["db_clt"]:
        oracledb.init_oracle_client(lib_dir=db_config["db_clt"])
        
    connection = oracledb.connect(
        f'{db_config["db_user"]}/{db_config["db_pswd"]}@{db_config["db_ip"]}:{db_config["db_port"]}/{db_config["db_name"]}'
    )
    return connection

def revenda(cnpj):
    db_config = load_db_config()
    conn = connect_to_db(db_config)
    cursor = conn.cursor()
    query = "SELECT empresa, revenda FROM ger_revenda WHERE cnpj = :cnpj"
    cursor.execute(query, [cnpj])
    result = cursor.fetchone()
    cursor.close()
    conn.close()
    return result

def cnpj(num_endereco, cidade):
    db_config = load_db_config()
    conn = connect_to_db(db_config)
    cursor = conn.cursor()
    query = "SELECT emp.cnpj FROM ger_revenda emp WHERE emp.NRO_ENDERECO = :num_endereco AND emp.CIDADE = :cidade"
    logging.info(f"Executando consulta: {query} com parâmetros: num_endereco={num_endereco}, cidade={cidade}")
    cursor.execute(query, [num_endereco, cidade])
    result = cursor.fetchone()
    logging.info(f"Resultado da consulta: {result}")
    cursor.close()
    conn.close()
    return result[0] if result else None

def login (sender):
    db_config = load_db_config()
    conn = connect_to_db(db_config)
    cursor = conn.cursor()
    query = "SELECT LOGIN FROM GER_USUARIO U WHERE U.EMAIL = :sender"  
    cursor.execute(query, [sender])
    result = cursor.fetchone()
    cursor.close()
    conn.close()
    return result