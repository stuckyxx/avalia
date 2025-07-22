import streamlit as st
import json
import os
from datetime import datetime, timedelta
import docx
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
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

def on_disponibilidade_change(secao, criterio, subcriterios):
    """Função para atualizar os subcritérios quando a Disponibilidade muda."""
    chave_disponibilidade = f"{secao}_{criterio}_Disponibilidade"
    novo_status_disponibilidade = st.session_state[chave_disponibilidade]
    st.session_state.respostas[chave_disponibilidade] = novo_status_disponibilidade
    novo_status_subs = "Atende" if novo_status_disponibilidade == "Atende" else "Não Atende"
    for sub in subcriterios:
        if sub != "Disponibilidade": st.session_state.respostas[f"{secao}_{criterio}_{sub}"] = novo_status_subs

# --- FUNÇÃO PRINCIPAL DE GERAÇÃO DE RELATÓRIO ---
def gerar_relatorio_pdf_paisagem(respostas, municipio, segmento, matriz_perguntas, tipo_relatorio, nome_usuario):
    """Gera o relatório em PDF com layout de tabela em modo paisagem."""
    doc = docx.Document()
    section = doc.sections[0]
    nova_largura, nova_altura = section.page_height, section.page_width
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = nova_largura
    section.page_height = nova_altura
    section.left_margin = Inches(0.5); section.right_margin = Inches(0.5)
    section.top_margin = Inches(0.75); section.bottom_margin = Inches(0.75)

    doc.add_heading(f'Relatório de Avaliação de Transparência - {segmento}', level=0)
    doc.add_heading(municipio, level=1)
    doc.add_paragraph(f"Data de Geração: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
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
        doc.add_paragraph("Nenhuma não conformidade foi encontrada para este tipo de relatório.", style='Body Text')
    else:
        for secao, perguntas in matriz_a_usar.items():
            if secao == "Municipios_MA": continue
            doc.add_heading(secao, level=2)
            num_cols = 4 if tipo_relatorio == "Apenas Não Conformidades" else 5
            table = doc.add_table(rows=1, cols=num_cols)
            table.style = 'Table Grid'
            table_xml = table._element; tblPr = table_xml.xpath('w:tblPr')[0]
            tbl_w = parse_xml(r'<w:tblW {} w:type="pct" w:w="5000"/>'.format(nsdecls('w')))
            tblPr.append(tbl_w)
            headers = ['Tópico', 'Matriz/Class.', 'Critério', 'Link / Observação'] if num_cols == 4 else ['Tópico', 'Matriz/Class.', 'Critério', 'Link / Observação', 'Atende?']
            hdr_cells = table.rows[0].cells
            for i, header_text in enumerate(headers): hdr_cells[i].text = header_text; hdr_cells[i].paragraphs[0].runs[0].font.bold = True
            for item in perguntas:
                row_cells = table.add_row().cells
                row_cells[0].text = item.get('topico', ''); row_cells[1].text = f"{item.get('matriz', '')}\n{item.get('classificacao', '')}"
                row_cells[2].paragraphs[0].add_run(item.get('criterio', ''))
                observacoes_finais, links_finais, status_geral_atende = [], [], True
                for subcriterio in item["subcriterios"]:
                    chave_resposta = f"{secao}_{item['criterio']}_{subcriterio}"
                    if respostas.get(chave_resposta) == "Não Atende":
                        status_geral_atende = False; obs = respostas.get(f"{chave_resposta}_obs", "")
                        if obs: observacoes_finais.append(f"{subcriterio}: {obs}")
                chave_links_pergunta = f"{secao}_{item['criterio']}_links"; links_finais = respostas.get(chave_links_pergunta, [])
                if num_cols == 5:
                    p_status = row_cells[2].add_paragraph(); run_status_label = p_status.add_run("\nStatus: "); run_status_label.font.bold = True
                    run_status = p_status.add_run("Atende") if status_geral_atende else p_status.add_run("Não Atende"); run_status.font.bold = True
                    run_status.font.color.rgb = RGBColor(0x00, 0x80, 0x00) if status_geral_atende else RGBColor(0xFF, 0x00, 0x00)
                p_obs_link = row_cells[3].paragraphs[0]; p_obs_link.text = ""
                if links_finais:
                    p_obs_link.add_run("Links:\n").bold = True
                    for link in list(set(links_finais)): p_obs_link.add_run(f"{link}\n")
                if observacoes_finais:
                    if links_finais: p_obs_link.add_run("\n")
                    p_obs_link.add_run("Observações:\n").bold = True
                    for obs_text in observacoes_finais: p_obs_link.add_run(f"{obs_text}\n")
                if num_cols == 5:
                    p_atende = row_cells[4].paragraphs[0]; p_atende.alignment = WD_ALIGN_PARAGRAPH.CENTER; p_atende.add_run("☑ ")
                    run_status = p_atende.add_run("✓") if status_geral_atende else p_atende.add_run("X")
                    run_status.font.color.rgb = RGBColor(0x00, 0x80, 0x00) if status_geral_atende else RGBColor(0xFF, 0x00, 0x00)
            doc.add_paragraph()
    doc.add_page_break()
    doc.add_heading('Resultado da Avaliação', level=2)
    resultados = calcular_indice_e_selo(respostas, matriz_perguntas)
    total_verificacoes, total_sim, total_nao = 0, 0, 0
    for secao, perguntas in matriz_perguntas.items():
        if secao == "Municipios_MA": continue
        for item in perguntas:
            total_verificacoes += 1
            if not any(respostas.get(f"{secao}_{item['criterio']}_{sub}") == "Não Atende" for sub in item["subcriterios"]): total_sim += 1
            else: total_nao += 1
    p_resultado = doc.add_paragraph(); p_resultado.add_run("Índice de Transparência: ").bold = True; p_resultado.add_run(f"{resultados['indice']:.2f}%")
    p_essenciais = doc.add_paragraph(); p_essenciais.add_run("Atendimento dos Critérios Essenciais: ").bold = True; p_essenciais.add_run(f"{resultados['percentual_essenciais']:.2f}%")
    if resultados['percentual_essenciais'] < 100: p_essenciais.add_run(" (Não qualificado para os selos Prata, Ouro ou Diamante)").italic = True
    p_selo = doc.add_paragraph(); p_selo.add_run("Selo Alcançado: ").bold = True; run_selo = p_selo.add_run(f"{resultados['selo']}"); run_selo.font.size = Pt(14); run_selo.font.bold = True
    doc.add_paragraph(); doc.add_heading('Resumo Detalhado', level=3)
    doc.add_paragraph(f"TOTAL DE CRITÉRIOS AVALIADOS: {total_verificacoes}\nCritérios Atendidos (Sim): {total_sim}\nCritérios Não Atendidos (Não): {total_nao}")
    doc.add_paragraph(f"USUÁRIO: {nome_usuario} - DATA/HORA DA VERIFICAÇÃO: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    prefixo = "NaoConformidades" if tipo_relatorio == "Apenas Não Conformidades" else "Completo"
    nome_base = f"Relatorio_{prefixo}_{segmento.replace(' ', '')}_{municipio.replace(' ', '')}_{timestamp}"
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
                                    chave_obs = f"{chave_resposta}_obs"
                                    # #############################################################
                                    # CORREÇÃO APLICADA AQUI
                                    # #############################################################
                                    obs = st.text_area("Observação:", value=st.session_state.respostas.get(chave_obs, ""), key=chave_obs)
                                    st.session_state.respostas[chave_obs] = obs
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
                                            chave_obs = f"{chave_resposta}_obs"
                                            # #############################################################
                                            # CORREÇÃO APLICADA AQUI
                                            # #############################################################
                                            obs = st.text_area("Observação:", value=st.session_state.respostas.get(chave_obs, ""), key=chave_obs)
                                            st.session_state.respostas[chave_obs] = obs
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
                    path_pdf = gerar_relatorio_pdf_paisagem(st.session_state.respostas, st.session_state.municipio, st.session_state.segmento, matriz_completa[st.session_state.segmento], tipo_relatorio, st.session_state["name"])
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