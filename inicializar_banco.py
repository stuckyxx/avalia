import streamlit as st
import os
import sqlite3
import psycopg2

# @st.cache_resource é a forma moderna e correta de gerenciar conexões.
# A conexão é criada uma vez e reutilizada, tornando o app muito mais rápido.
@st.cache_resource
def get_db_connection():
    """
    Cria e retorna uma conexão com o banco de dados apropriado (SQLite ou PostgreSQL)
    lendo a configuração do st.secrets.
    """
    # Verifica se a configuração de produção (nuvem) existe nas secrets
    if "database_prod" in st.secrets:
        config = st.secrets["database_prod"]
        db_type = config["type"]
    else: # Senão, usa a configuração local
        config = st.secrets["database_local"]
        db_type = config["type"]

    try:
        if db_type == "sqlite":
            db_path = config["path"]
            os.makedirs(os.path.dirname(db_path), exist_ok=True)
            # check_same_thread=False é crucial para o Streamlit com SQLite
            conn = sqlite3.connect(db_path, check_same_thread=False)
            st.toast("Conectado ao banco de dados local (SQLite).")
        
        elif db_type == "postgresql":
            conn = psycopg2.connect(config["url"])
            st.toast("Conectado ao banco de dados da nuvem (PostgreSQL).")
        
        else:
            st.error(f"Tipo de banco de dados '{db_type}' não suportado.")
            return None
        
        return conn

    except (sqlite3.Error, psycopg2.Error) as e:
        st.error(f"Erro ao conectar ao banco de dados ({db_type}): {e}")
        return None

def initialize_database(_conn):
    """
    Recebe uma conexão e garante que as tabelas existam, usando a sintaxe SQL correta.
    """
    db_type = st.secrets.get("database_prod", st.secrets["database_local"])["type"]

    if db_type == "sqlite":
        sql_create_avaliacoes = """
            CREATE TABLE IF NOT EXISTS avaliacoes (...);
        """ # Sintaxe SQLite (id INTEGER PRIMARY KEY AUTOINCREMENT)
        sql_create_respostas = """
            CREATE TABLE IF NOT EXISTS respostas (...);
        """
    elif db_type == "postgresql":
        sql_create_avaliacoes = """
            CREATE TABLE IF NOT EXISTS avaliacoes (
                id SERIAL PRIMARY KEY,
                municipio VARCHAR(255) NOT NULL,
                segmento VARCHAR(255) NOT NULL,
                usuario VARCHAR(255) NOT NULL,
                data_inicio TIMESTAMPTZ DEFAULT NOW(),
                status VARCHAR(50) DEFAULT 'Em Andamento',
                indice_final REAL
            );
        """
        sql_create_respostas = """
            CREATE TABLE IF NOT EXISTS respostas (
                id SERIAL PRIMARY KEY,
                avaliacao_id INTEGER NOT NULL,
                chave VARCHAR(255) NOT NULL,
                valor TEXT,
                UNIQUE(avaliacao_id, chave),
                FOREIGN KEY (avaliacao_id) REFERENCES avaliacoes (id)
            );
        """
    
    try:
        with _conn.cursor() as cursor:
            cursor.execute(sql_create_avaliacoes)
            cursor.execute(sql_create_respostas)
        _conn.commit()
    except (sqlite3.Error, psycopg2.Error) as e:
        st.error(f"Erro ao inicializar tabelas: {e}")
