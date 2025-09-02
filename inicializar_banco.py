# Arquivo: inicializar_banco.py

import sqlite3
import os
import streamlit as st

DB_FILE = "data/database.db"

def conectar_db():
    """Cria e retorna uma conexão com o banco de dados SQLite."""
    try:
        os.makedirs("data", exist_ok=True)
        conn = sqlite3.connect(DB_FILE)
        return conn
    except sqlite3.Error as e:
        st.error(f"Erro ao conectar ao banco de dados: {e}")
        return None

def inicializar():
    """Garante que o banco de dados e suas tabelas existam."""
    conn = conectar_db()
    if conn is not None:
        try:
            cursor = conn.cursor()
            # Tabela para guardar o resumo de cada avaliação
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS avaliacoes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    municipio TEXT NOT NULL,
                    segmento TEXT NOT NULL,
                    usuario TEXT NOT NULL,
                    data_inicio TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    status TEXT DEFAULT 'Em Andamento',
                    indice_final REAL
                );
            """)
            # Tabela para guardar cada resposta individual
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS respostas (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    avaliacao_id INTEGER NOT NULL,
                    chave TEXT NOT NULL,
                    valor TEXT,
                    UNIQUE(avaliacao_id, chave),
                    FOREIGN KEY (avaliacao_id) REFERENCES avaliacoes (id)
                );
            """)
            conn.commit()
        except sqlite3.Error as e:
            st.error(f"Erro ao criar tabelas: {e}")
        finally:
            conn.close()