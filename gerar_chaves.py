# --- gerar_chaves.py ---
import streamlit_authenticator as stauth

# Lista de senhas em texto simples que você quer criptografar
senhas = ['admin123', 'senha456'] 

hashed_passwords = stauth.Hasher(senhas).generate()
print(hashed_passwords)