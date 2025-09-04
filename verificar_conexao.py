# Arquivo: verificar_conexao.py
import psycopg2
import sys

# --- ATENÇÃO AQUI ---
# Cole a URL do "Transaction pooler" do Supabase abaixo.
# Substitua 'SUA_NOVA_SENHA_SIMPLES_AQUI' pela senha que você acabou de criar (ex: 'avaliapntp2025').
DB_URL = "postgresql://postgres.ikdckhgxabiiufnbmbbv:SUA_NOVA_SENHA_SIMPLES_AQUI@aws-1-sa-east-1.pooler.supabase.com:6543/postgres"

print("Tentando conectar ao banco de dados...")

try:
    conn = psycopg2.connect(DB_URL)
    print("\n✅ SUCESSO! A conexão com o banco de dados foi estabelecida.")
    print("   A URL e a SENHA estão PERFEITAS.")
    conn.close()
except Exception as e:
    print("\n❌ FALHA! A conexão com o banco de dados falhou.")
    print("   O problema está na URL ou na senha. Verifique os dados e tente novamente.")
    print("\n   Detalhe do erro:")
    print(e)
    sys.exit(1)
