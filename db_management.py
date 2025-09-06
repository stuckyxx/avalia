# Arquivo: db_management.py

import sqlite3
import os
import streamlit as st

DB_PATH = "data/database.db"

@st.cache_resource
def get_db_connection():
    """
    Cria e gerencia uma conexão única com o banco de dados SQLite local.
    """
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    return conn

def initialize_database(_conn):
    """Garante que as tabelas existam no banco de dados SQLite."""
    sql_create_avaliacoes = """
        CREATE TABLE IF NOT EXISTS avaliacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            municipio TEXT NOT NULL,
            segmento TEXT NOT NULL,
            usuario TEXT NOT NULL,
            data_inicio TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            status TEXT DEFAULT 'Em Andamento',
            indice_final REAL
        );
    """
    sql_create_respostas = """
        CREATE TABLE IF NOT EXISTS respostas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            avaliacao_id INTEGER NOT NULL,
            chave TEXT NOT NULL,
            valor TEXT,
            UNIQUE(avaliacao_id, chave),
            FOREIGN KEY (avaliacao_id) REFERENCES avaliacoes (id)
        );
    """
    try:
        cursor = _conn.cursor()
        cursor.execute(sql_create_avaliacoes)
        cursor.execute(sql_create_respostas)
        _conn.commit()
    except sqlite3.Error as e:
        st.error(f"Erro ao inicializar tabelas: {e}")
