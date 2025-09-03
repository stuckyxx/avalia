import streamlit as st
import os
import sqlite3
import psycopg2

# @st.cache_resource é a forma correta de gerenciar conexões em produção.
@st.cache_resource
def get_db_connection():
    """
    Lê a configuração do st.secrets e conecta ao banco de dados apropriado.
    A conexão é criada uma vez e reutilizada, tornando o app muito mais rápido.
    """
    # No Streamlit Cloud, as secrets de produção serão encontradas.
    if "database_prod" in st.secrets:
        config = st.secrets["database_prod"]
    # Localmente, ele usará esta seção se a de produção não existir.
    else:
        config = st.secrets["database_local"]
    
    db_type = config["type"]

    try:
        if db_type == "sqlite":
            db_path = config["path"]
            os.makedirs(os.path.dirname(db_path), exist_ok=True)
            conn = sqlite3.connect(db_path, check_same_thread=False)
        elif db_type == "postgresql":
            conn = psycopg2.connect(config["url"])
        else:
            st.error(f"Tipo de banco de dados '{db_type}' não suportado.")
            return None
        return conn
    except (sqlite3.Error, psycopg2.Error) as e:
        st.error(f"Erro ao conectar ao banco de dados ({db_type}): {e}")
        return None

def initialize_database(_conn):
    """Recebe uma conexão e garante que as tabelas existam com a sintaxe SQL correta."""
    db_type = st.secrets.get("database_prod", st.secrets["database_local"])["type"]

    # Sintaxe SQL para PostgreSQL (usada na nuvem)
    sql_create_avaliacoes_pg = """
        CREATE TABLE IF NOT EXISTS avaliacoes (
            id SERIAL PRIMARY KEY, municipio VARCHAR(255) NOT NULL, segmento VARCHAR(255) NOT NULL,
            usuario VARCHAR(255) NOT NULL, data_inicio TIMESTAMPTZ DEFAULT NOW(),
            status VARCHAR(50) DEFAULT 'Em Andamento', indice_final REAL
        );"""
    sql_create_respostas_pg = """
        CREATE TABLE IF NOT EXISTS respostas (
            id SERIAL PRIMARY KEY, avaliacao_id INTEGER NOT NULL, chave VARCHAR(255) NOT NULL,
            valor TEXT, UNIQUE(avaliacao_id, chave), FOREIGN KEY (avaliacao_id) REFERENCES avaliacoes (id)
        );"""

    # Sintaxe SQL para SQLite (usada localmente)
    sql_create_avaliacoes_sqlite = """
        CREATE TABLE IF NOT EXISTS avaliacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT, municipio TEXT NOT NULL, segmento TEXT NOT NULL,
            usuario TEXT NOT NULL, data_inicio TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            status TEXT DEFAULT 'Em Andamento', indice_final REAL
        );"""
    sql_create_respostas_sqlite = """
        CREATE TABLE IF NOT EXISTS respostas (
            id INTEGER PRIMARY KEY AUTOINCREMENT, avaliacao_id INTEGER NOT NULL, chave TEXT NOT NULL,
            valor TEXT, UNIQUE(avaliacao_id, chave), FOREIGN KEY (avaliacao_id) REFERENCES avaliacoes (id)
        );"""

    try:
        with _conn.cursor() as cursor:
            if db_type == "postgresql":
                cursor.execute(sql_create_avaliacoes_pg)
                cursor.execute(sql_create_respostas_pg)
            else: # sqlite
                cursor.execute(sql_create_avaliacoes_sqlite)
                cursor.execute(sql_create_respostas_sqlite)
        _conn.commit()
    except (sqlite3.Error, psycopg2.Error) as e:
        st.error(f"Erro ao inicializar tabelas: {e}")
