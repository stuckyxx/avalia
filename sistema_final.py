import streamlit as st
import json
import os
from datetime import datetime, timedelta
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx2pdf import convert
import streamlit_authenticator as stauth
import yaml
from yaml.loader import SafeLoader

# --- FUNÇÕES AUXILIARES ---
@st.cache_data
def carregar_criterios_do_arquivo(caminho_arquivo="criterios_por_topico.json"):
    """Carrega os critérios de avaliação e a lista de municípios do arquivo JSON."""
    try:
        with open(caminho_arquivo, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        st.error(f"ERRO: O arquivo de dados '{caminho_arquivo}' não foi encontrado.")
        return None
    except json.JSONDecodeError:
        st.error(f"ERRO: O arquivo '{caminho_arquivo}' contém um erro de formatação.")
        return None

def criar_pastas_necessarias():
    """Cria as pastas para salvar os dados e relatórios."""
    os.makedirs("data/avaliacoes", exist_ok=True)
    os.makedirs("relatorios", exist_ok=True)

def calcular_indice_e_selo(respostas, matriz_perguntas):
    """Calcula o índice de transparência e o selo Atricon com base nos pesos."""
    pesos = {"ESSENCIAL": 2.0, "OBRIGATÓRIA": 1.5, "RECOMENDADA": 1.0}
    total_pontos_possiveis, pontos_obtidos, total_essenciais, essenciais_atendidos = 0, 0, 0, 0
    for secao, perguntas in matriz_perguntas.items():
        if secao == "Municipios_MA": continue
        for item in perguntas:
            classificacao = item.get("classificacao", "RECOMENDADA").upper()
            peso = pesos.get(classificacao, 1.0)
            total_pontos_possiveis += peso
            status_geral_atende = not any(respostas.get(f"{secao}_{item['criterio']}_{sub}") == "Não Atende" for sub in item["subcriterios"])
            if status_geral_atende: pontos_obtidos += peso
            if classificacao == "ESSENCIAL":
                total_essenciais += 1
                if status_geral_atende: essenciais_atendidos += 1
    percentual_essenciais = (essenciais_atendidos / total_essenciais * 100) if total_essenciais > 0 else 100
    indice = (pontos_obtidos / total_pontos_possiveis * 100) if total_pontos_possiveis > 0 else 0
    selo = "Inexistente"
    if indice > 0:
        if percentual_essenciais == 100:
            if indice >= 95: selo = "💎 Diamante"
            elif indice >= 85: selo = "🥇 Ouro"
            elif indice >= 75: selo = "🥈 Prata"
            else: selo = "Elevado (não elegível para selo)"
        else:
            if indice >= 75: selo = "Elevado"
            elif indice >= 50: selo = "Intermediário"
            elif indice >= 30: selo = "Básico"
            else: selo = "Inicial"
    return {"indice": indice, "selo": selo, "percentual_essenciais": percentual_essenciais}

def calcular_pontuacao_secao(respostas, perguntas_secao, nome_secao):
    """Calcula a pontuação de uma seção específica."""
    pesos = {"ESSENCIAL": 2.0, "OBRIGATÓRIA": 1.5, "RECOMENDADA": 1.0}
    total_pontos_possiveis, pontos_obtidos = 0, 0
    for item in perguntas_secao:
        classificacao = item.get("classificacao", "RECOMENDADA").upper()
        peso = pesos.get(classificacao, 1.0)
        total_pontos_possiveis += peso
        if not any(respostas.get(f"{nome_secao}_{item['criterio']}_{sub}") == "Não Atende" for sub in item["subcriterios"]):
            pontos_obtidos += peso
    return (pontos_obtidos / total_pontos_possiveis * 100) if total_pontos_possiveis > 0 else 100

def on_disponibilidade_change(secao, criterio, subcriterios):
    chave_disponibilidade = f"{secao}_{criterio}_Disponibilidade"
    novo_status_disponibilidade = st.session_state[chave_disponibilidade]
    st.session_state.respostas[chave_disponibilidade] = novo_status_disponibilidade
    novo_status_subs = "Atende" if novo_status_disponibilidade == "Atende" else "Não Atende"
    for sub in subcriterios:
        if sub != "Disponibilidade": st.session_state.respostas[f"{secao}_{criterio}_{sub}"] = novo_status_subs

# --- FUNÇÃO DE GERAÇÃO DE RELATÓRIO (COM AJUSTES DE LAYOUT) ---
def gerar_relatorio_novo_modelo(respostas, municipio, segmento, matriz_perguntas, tipo_relatorio, nome_usuario, usuario_config):
    template_tipo = usuario_config.get('template', 'padrao')
    template_path = f"modelo_{template_tipo}.docx"
    try:
        doc = docx.Document(template_path)
    except Exception:
        st.sidebar.error(f"ERRO: Arquivo de modelo '{template_path}' não foi encontrado."); return None
    
    # --- PÁGINA DE ROSTO ---
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run("PADRÃO MÍNIMO DE QUALIDADE")
    run_title.font.size = Pt(22); run_title.bold = True
    doc.add_paragraph() #paragrafo

    p_subtitulo = doc.add_paragraph(); p_subtitulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_subtitulo.add_run("Relatório de Transparência\n").bold = True
    doc.add_paragraph()
    p_subtitulo.add_run(f"{segmento} de {municipio}").bold = True
    
    resultados = calcular_indice_e_selo(respostas, matriz_perguntas)
    doc.add_paragraph()
    p_score = doc.add_paragraph(f"{resultados['indice']:.2f}%"); p_score.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_score.runs[0].font.size = Pt(48); p_score.runs[0].bold = True
    p_selo = doc.add_paragraph(f"{resultados['selo']}"); p_selo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_selo.runs[0].font.size = Pt(24); p_selo.runs[0].bold = True
    doc.add_paragraph()
    doc.add_paragraph() #paragrafo
    
    texto_intro = f"Com base na Lei 12.527/2011 (Lei de Acesso à Informação), o nosso controle de qualidade fez uma avaliação geral da {segmento} de {municipio}, na qual, apresentou as seguintes informações:"
    doc.add_paragraph(texto_intro); doc.add_paragraph()
    doc.add_paragraph(f"Exercício: {datetime.now().year}")
    doc.add_paragraph() #paragrafo
    doc.add_paragraph(f"Avaliação feita por: {nome_usuario}")
    doc.add_paragraph() #paragrafo
    doc.add_paragraph(f"Data de Geração: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    
    # --- PÁGINAS DE DETALHAMENTO ---
    doc.add_page_break()
    p_detalhe = doc.add_paragraph()
    run_detalhe = p_detalhe.add_run("Detalhamento da Avaliação")
    run_detalhe.font.size = Pt(18); run_detalhe.bold = True
    doc.add_paragraph()
    
    matriz_a_usar = matriz_perguntas
    if tipo_relatorio == "Apenas Não Conformidades":
        perguntas_filtradas = {}
        for secao, perguntas in matriz_perguntas.items():
            if secao == "Municipios_MA": continue
            itens_nao_conformes = [item for item in perguntas if any(respostas.get(f"{secao}_{item['criterio']}_{sub}") == "Não Atende" for sub in item["subcriterios"])]
            if itens_nao_conformes: perguntas_filtradas[secao] = itens_nao_conformes
        matriz_a_usar = perguntas_filtradas

    if not matriz_a_usar:
        doc.add_paragraph("Nenhuma não conformidade foi encontrada.")
    else:
        for secao, perguntas in matriz_a_usar.items():
            if secao == "Municipios_MA": continue
            score_secao = calcular_pontuacao_secao(respostas, matriz_perguntas[secao], secao)
            p_secao_titulo = doc.add_paragraph()
            run_secao_titulo = p_secao_titulo.add_run(f"{secao.upper()} - {score_secao:.2f}%")
            run_secao_titulo.font.size = Pt(14); run_secao_titulo.bold = True
            doc.add_paragraph()

            for item in perguntas:
                p_item = doc.add_paragraph()
                p_item.add_run(f"{item['topico']} - {item['criterio']} ({item.get('classificacao', '').upper()})").bold = True
                doc.add_paragraph() #paragrafo
                observacoes_finais = []
                subcriterios_nao_atendidos = []
                
                for subcriterio in item["subcriterios"]:
                    chave_resposta = f"{secao}_{item['criterio']}_{subcriterio}"
                    if respostas.get(chave_resposta) == "Não Atende":
                        subcriterios_nao_atendidos.append(subcriterio)
                        obs = respostas.get(f"{chave_resposta}_obs", "")
                        if obs: observacoes_finais.append((subcriterio, obs))

                if subcriterios_nao_atendidos:
                    for sub in subcriterios_nao_atendidos:
                        p_item.add_run(f"\n• {sub}: ").italic = True
                        run_status = p_item.add_run("Não Atende"); run_status.bold = True; run_status.font.color.rgb = RGBColor(0xFF, 0, 0)
                elif tipo_relatorio == "Relatório Completo":
                     p_item.add_run(f"\n• Status: ").italic = True
                     run_status = p_item.add_run("Atende"); run_status.bold = True; run_status.font.color.rgb = RGBColor(0x00, 0x80, 0x00)

                # --- [INÍCIO DA ALTERAÇÃO] ---
                chave_links_pergunta = f"{secao}_{item['criterio']}_links"
                links_finais = list(set(respostas.get(chave_links_pergunta, [])))

                if links_finais or observacoes_finais:
                    # Adiciona o cabeçalho principal
                    doc.add_paragraph("Links e Observações:").runs[0].bold = True

                    # Adiciona os links em um parágrafo próprio
                    if links_finais:
                        p_links = doc.add_paragraph()
                        # Constrói o texto dos links
                        texto_links = "\n".join([f"- Link: {link}" for link in links_finais])
                        p_links.add_run(texto_links)
                    
                    # Se ambos existirem, adiciona o separador em seu próprio parágrafo
                    if links_finais and observacoes_finais:
                        doc.add_paragraph("")

                    # Adiciona as observações em um parágrafo próprio
                    if observacoes_finais:
                        p_obs = doc.add_paragraph()
                        for i, (sub, obs_text) in enumerate(observacoes_finais):
                            if i > 0: p_obs.add_run("\n")
                            p_obs.add_run(f"- Observação ({sub}): ")
                            run_obs = p_obs.add_run(obs_text)
                            run_obs.font.color.rgb = RGBColor(0xFF, 0, 0)
                # --- [FIM DA ALTERAÇÃO] ---

            doc.add_paragraph()

    # --- SALVAMENTO E CONVERSÃO ---
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_base = f"Relatorio_Final_{segmento.replace(' ', '')}_{municipio.replace(' ', '')}_{timestamp}"
    path_docx = os.path.join("relatorios", f"{nome_base}.docx"); path_pdf = os.path.join("relatorios", f"{nome_base}.pdf")
    doc.save(path_docx)
    try:
        convert(path_docx, path_pdf); os.remove(path_docx)
        return path_pdf
    except Exception as e:
        st.sidebar.error(f"Falha ao converter para PDF: {e}"); st.session_state.fallback_docx_path = path_docx
        return None

# --- INTERFACE GRÁFICA ---
st.set_page_config(layout="wide", page_title="Avaliador de Transparência")
st.title("📄 Sistema de Avaliação de Transparência Municipal")
matriz_completa = carregar_criterios_do_arquivo()

if matriz_completa:
    try:
        with open('config.yaml', 'r', encoding='utf-8') as file: config = yaml.load(file, Loader=SafeLoader)
        authenticator = stauth.Authenticate(config['credentials'], config['cookie']['name'], config['cookie']['key'], config['cookie']['expiry_days'])
        authenticator.login('main')
    except FileNotFoundError:
        st.error("ERRO: O arquivo 'config.yaml' não foi encontrado."); st.stop()
    if st.session_state["authentication_status"]:
        authenticator.logout('Logout', 'sidebar', key='logout_button')
        st.sidebar.title(f"Bem-vindo(a),\n{st.session_state['name']}!")
        criar_pastas_necessarias()
        st.sidebar.header("Configuração da Avaliação")
        
        MUNICIPIOS_MARANHAO = ["- Selecione um município -"] + sorted(matriz_completa.get("Municipios_MA", []))
        municipio = st.sidebar.selectbox("Nome do Município", options=MUNICIPIOS_MARANHAO)
        
        opcoes_segmento = [key for key in matriz_completa.keys() if key != "Municipios_MA"]
        segmento = st.sidebar.selectbox("Órgão/Poder", opcoes_segmento)
        
        if municipio != "- Selecione um município -" and segmento:
            nome_arquivo_avaliacao = f"avaliacao_{segmento.replace(' ', '')}_{municipio.replace(' ', '')}_{st.session_state['username']}.json"
            caminho_arquivo = os.path.join("data/avaliacoes", nome_arquivo_avaliacao)
            if st.sidebar.button("✅ Iniciar / Continuar Avaliação"):
                if os.path.exists(caminho_arquivo):
                    with open(caminho_arquivo, 'r', encoding='utf-8') as f: st.session_state.respostas = json.load(f)
                    st.sidebar.success("Avaliação anterior carregada!")
                else:
                    st.session_state.respostas = {}; st.sidebar.info("Iniciando uma nova avaliação.")
                st.session_state.path_pdf = None; st.session_state.fallback_docx_path = None
                st.session_state.avaliacao_iniciada = True; st.session_state.caminho_arquivo = caminho_arquivo
                st.session_state.municipio = municipio; st.session_state.segmento = segmento
                st.session_state.last_save_time = datetime.now()
                
        if st.session_state.get('avaliacao_iniciada', False):
            if 'last_save_time' not in st.session_state: st.session_state.last_save_time = datetime.now()
            if datetime.now() - st.session_state.last_save_time > timedelta(minutes=10):
                try:
                    with open(st.session_state.caminho_arquivo, 'w', encoding='utf-8') as f: json.dump(st.session_state.respostas, f, ensure_ascii=False, indent=4)
                    st.session_state.last_save_time = datetime.now()
                    st.toast(f"Progresso salvo automaticamente às {datetime.now().strftime('%H:%M:%S')}")
                except Exception as e: st.toast(f"Erro no salvamento automático: {e}")
            st.header(f"Avaliação: {st.session_state.municipio} - {st.session_state.segmento}")
            matriz_perguntas = matriz_completa[st.session_state.segmento]
            for secao, perguntas in matriz_perguntas.items():
                with st.expander(f"**{secao}**", expanded=False):
                    for item in perguntas:
                        st.markdown(f"#### {item['topico']} - {item['criterio']}"); st.markdown("---")
                        col_link_ui, _ = st.columns([1, 1])
                        with col_link_ui:
                            st.subheader("Links de Evidência")
                            chave_links = f"{secao}_{item['criterio']}_links"
                            if chave_links not in st.session_state.respostas: st.session_state.respostas[chave_links] = []
                            for i, link in enumerate(st.session_state.respostas[chave_links]):
                                link_cols = st.columns([10, 1]); link_cols[0].info(link)
                                if link_cols[1].button("X", key=f"rem_{chave_links}_{i}"): st.session_state.respostas[chave_links].pop(i); st.rerun()
                            link_cols = st.columns([10, 1])
                            novo_link = link_cols[0].text_input("Adicionar novo link", key=f"add_{chave_links}", label_visibility="collapsed")
                            if link_cols[1].button("➕", key=f"btn_{chave_links}"):
                                if novo_link: st.session_state.respostas[chave_links].append(novo_link); st.rerun()
                        st.markdown("---"); st.subheader("Critérios de Avaliação")
                        subcriterios = item["subcriterios"]
                        if "Disponibilidade" in subcriterios:
                            cols = st.columns([1, 2]); subcriterio = "Disponibilidade"; chave_resposta = f"{secao}_{item['criterio']}_{subcriterio}"
                            with cols[0]:
                                resposta_atual = st.session_state.respostas.get(chave_resposta, "Atende")
                                st.radio(subcriterio, ("Atende", "Não Atende"), index=1 if resposta_atual == "Não Atende" else 0, key=chave_resposta, horizontal=True, on_change=on_disponibilidade_change, kwargs=dict(secao=secao, criterio=item['criterio'], subcriterios=subcriterios))
                            if st.session_state.respostas.get(chave_resposta) == "Não Atende":
                                with cols[1]:
                                    chave_obs = f"{chave_resposta}_obs"; obs = st.text_area("Observação:", value=st.session_state.respostas.get(chave_obs, ""), key=chave_obs); st.session_state.respostas[chave_obs] = obs
                        chave_disponibilidade = f"{secao}_{item['criterio']}_Disponibilidade"
                        disponibilidade_falhou = ("Disponibilidade" in subcriterios and st.session_state.respostas.get(chave_disponibilidade) == "Não Atende")
                        if not disponibilidade_falhou:
                            for subcriterio in subcriterios:
                                if subcriterio != "Disponibilidade":
                                    cols = st.columns([1, 2]); chave_resposta = f"{secao}_{item['criterio']}_{subcriterio}"
                                    with cols[0]:
                                        resposta_atual = st.session_state.respostas.get(chave_resposta, "Atende"); resposta = st.radio(subcriterio, ("Atende", "Não Atende"), index=1 if resposta_atual == "Não Atende" else 0, key=chave_resposta, horizontal=True); st.session_state.respostas[chave_resposta] = resposta
                                    if resposta == "Não Atende":
                                        with cols[1]:
                                            chave_obs = f"{chave_resposta}_obs"; obs = st.text_area("Observação:", value=st.session_state.respostas.get(chave_obs, ""), key=chave_obs); st.session_state.respostas[chave_obs] = obs
                        st.markdown("---")
            st.sidebar.header("Ações")
            if st.sidebar.button("💾 Salvar Progresso"):
                with open(st.session_state.caminho_arquivo, 'w', encoding='utf-8') as f: json.dump(st.session_state.respostas, f, ensure_ascii=False, indent=4)
                st.session_state.last_save_time = datetime.now(); st.sidebar.success("Progresso salvo!")
            st.sidebar.markdown("##### Tipo de Relatório"); tipo_relatorio = st.sidebar.radio("Escolha o tipo:", ("Apenas Não Conformidades", "Relatório Completo"), label_visibility="collapsed")
            if st.sidebar.button("📊 Gerar Relatório PDF"):
                with st.spinner("Gerando relatório PDF..."):
                    with open(st.session_state.caminho_arquivo, 'w', encoding='utf-8') as f: json.dump(st.session_state.respostas, f, ensure_ascii=False, indent=4)
                    st.session_state.fallback_docx_path = None
                    path_pdf = gerar_relatorio_novo_modelo(st.session_state.respostas, st.session_state.municipio, st.session_state.segmento, matriz_completa[st.session_state.segmento], tipo_relatorio, st.session_state["name"], config['credentials']['usernames'][st.session_state['username']])
                    st.session_state.path_pdf = path_pdf
                if st.session_state.path_pdf: st.sidebar.success("Relatório PDF pronto!")
            if st.session_state.get('path_pdf'):
                with open(st.session_state.path_pdf, "rb") as pdf_file: st.sidebar.download_button(label="⬇️ Baixar Relatório (.pdf)", data=pdf_file, file_name=os.path.basename(st.session_state.path_pdf), mime="application/pdf", key="download_pdf")
            if st.session_state.get('fallback_docx_path'):
                with open(st.session_state.fallback_docx_path, "rb") as docx_file: st.sidebar.download_button(label="⬇️ Baixar Arquivo Word (.docx)", data=docx_file, file_name=os.path.basename(st.session_state.fallback_docx_path), mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="download_docx_fallback")
    elif st.session_state["authentication_status"] is False: st.error('Usuário ou senha incorretos.')
    elif st.session_state["authentication_status"] is None: st.warning('Por favor, insira seu usuário e senha.')
else:
    st.warning("Aguardando o carregamento do arquivo 'criterios_por_topico.json'...")
