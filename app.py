import streamlit as st
import os
import json
from io import BytesIO
from datetime import datetime
from docxtpl import DocxTemplate, RichText

st.set_page_config(
    page_title="Proanalisis v1.3",
    page_icon="📐",
    layout="wide"
)

# CSS personalizado para fundo azul ergonômico com melhor contraste
st.markdown("""
<style>
    /* Fundo principal da aplicação */
    .stApp {
        background: linear-gradient(135deg, #e8f0fe 0%, #d4e4fc 100%);
    }
    
    /* Fundo dos containers principais */
    .main > div {
        background-color: rgba(255, 255, 255, 0.95);
        border-radius: 12px;
        padding: 1rem;
    }
    
    /* Sidebar com fundo azul mais escuro */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #1a5276 0%, #1a3a5c 100%);
    }
    
    [data-testid="stSidebar"] * {
        color: white !important;
    }
    
    [data-testid="stSidebar"] .stSelectbox label,
    [data-testid="stSidebar"] .stRadio label {
        color: white !important;
    }
    
    [data-testid="stSidebar"] select {
        background-color: #2c5a7a !important;
        color: white !important;
    }
    
    [data-testid="stSidebar"] .stTextInput input {
        background-color: #2c5a7a !important;
        color: white !important;
    }
    
    [data-testid="stSidebar"] .stTextInput label {
        color: white !important;
    }
    
    /* Cards de inconformidades */
    .card {
        padding: 0.8rem 1rem;
        border: 1px solid #c5d5e6;
        border-radius: 10px;
        background: #f8fafd !important;
        margin-bottom: 0.6rem;
        color: #1a1a1a !important;
    }
    
    .card b, .card strong {
        color: #1a5276 !important;
    }
    
    /* CORREÇÃO DE CONTRASTE PARA CAMPOS DE INPUT */
    .stTextInput label, .stSelectbox label, .stTextArea label, .stNumberInput label {
        color: #1a5276 !important;
        font-weight: 600 !important;
        font-size: 14px !important;
    }
    
    .stTextInput input {
        background-color: white !important;
        color: #1a1a1a !important;
        border: 1px solid #c5d5e6 !important;
        border-radius: 6px !important;
        font-size: 14px !important;
    }
    
    .stTextInput input:focus {
        border-color: #1a5276 !important;
        box-shadow: 0 0 0 2px rgba(26, 82, 118, 0.2) !important;
    }
    
    .stTextArea textarea {
        background-color: white !important;
        color: #1a1a1a !important;
        border: 1px solid #c5d5e6 !important;
        border-radius: 6px !important;
        font-size: 14px !important;
    }
    
    .stTextArea textarea:focus {
        border-color: #1a5276 !important;
        box-shadow: 0 0 0 2px rgba(26, 82, 118, 0.2) !important;
    }
    
    .stNumberInput input {
        background-color: white !important;
        color: #1a1a1a !important;
        border: 1px solid #c5d5e6 !important;
        border-radius: 6px !important;
    }
    
    /* Select boxes - fundo claro e texto escuro */
    .stSelectbox select {
        background-color: white !important;
        color: #1a1a1a !important;
        border: 1px solid #c5d5e6 !important;
        border-radius: 6px !important;
        font-size: 14px !important;
        font-weight: normal !important;
    }
    
    .stSelectbox select:focus {
        border-color: #1a5276 !important;
        box-shadow: 0 0 0 2px rgba(26, 82, 118, 0.2) !important;
    }
    
    /* CORREÇÃO PARA O DROPDOWN DAS CAIXAS DE SELEÇÃO */
    div[data-baseweb="select"] > div {
        background-color: white !important;
        color: #1a1a1a !important;
    }
    
    [data-testid="stSelectbox"] div[role="listbox"] {
        background-color: white !important;
        border: 1px solid #c5d5e6 !important;
        border-radius: 6px !important;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1) !important;
    }
    
    [data-testid="stSelectbox"] div[role="option"] {
        background-color: white !important;
        color: #1a1a1a !important;
        padding: 8px 12px !important;
        font-size: 14px !important;
    }
    
    [data-testid="stSelectbox"] div[role="option"]:hover {
        background-color: #e8f0fe !important;
        color: #1a5276 !important;
    }
    
    [data-testid="stSelectbox"] div[role="option"][aria-selected="true"] {
        background-color: #1a5276 !important;
        color: white !important;
    }
    
    .stSelectbox [data-baseweb="select"] input {
        background-color: white !important;
        color: #1a1a1a !important;
        border: 1px solid #c5d5e6 !important;
        border-radius: 6px !important;
    }
    
    .stSelectbox [data-baseweb="select"] div[class*="placeholder"] {
        color: #888888 !important;
    }
    
    .stSelectbox [data-baseweb="select"] [data-testid="stMarkdownContainer"] {
        background-color: transparent !important;
    }
    
    .stSelectbox [data-baseweb="select"] span {
        color: #1a1a1a !important;
    }
    
    ::placeholder {
        color: #888888 !important;
        opacity: 1 !important;
    }
    
    .status-badge {
        padding: 8px 12px;
        border-radius: 6px;
        margin-top: 28px;
        text-align: center;
    }
    
    .status-icon {
        font-size: 20px;
        display: block;
    }
    
    .status-text {
        font-weight: 600;
        font-size: 0.85rem;
    }
    
    .progress-wrap {
        width: 100%;
        background: #e9ecef;
        border-radius: 999px;
        height: 14px;
        overflow: hidden;
        margin: 8px 0 4px 0;
    }
    
    .progress-bar {
        height: 14px;
        border-radius: 999px;
        transition: width 0.3s ease;
    }
    
    /* BOTÕES - FONTE BRANCA */
    .stButton > button {
        border-radius: 8px !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
        color: white !important;
        border: none !important;
        font-size: 14px !important;
    }
    
    .stButton > button[kind="primary"] {
        background-color: #0d6e2e !important;
        color: white !important;
    }
    
    .stButton > button[kind="primary"]:hover {
        background-color: #0f8a3a !important;
        color: white !important;
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    
    .stButton > button:not([kind="primary"]) {
        background-color: #1a5276 !important;
        color: white !important;
    }
    
    .stButton > button:not([kind="primary"]):hover {
        background-color: #2c6b96 !important;
        color: white !important;
        transform: translateY(-2px);
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    
    .stButton > button:disabled {
        background-color: #cccccc !important;
        color: #666666 !important;
        cursor: not-allowed !important;
    }
    
    button, button p {
        color: white !important;
    }
    
    .stDownloadButton button {
        background-color: #0d6e2e !important;
        color: white !important;
    }
    
    .stDownloadButton button:hover {
        background-color: #0f8a3a !important;
        color: white !important;
    }
    
    h1, h2, h3, h4 {
        color: #1a5276 !important;
        font-weight: 600 !important;
    }
    
    p, li, .stMarkdown, .stText {
        color: #2c3e50 !important;
    }
    
    .stInfo, .stSuccess, .stWarning, .stError {
        border-radius: 8px !important;
    }
    
    [data-testid="stMetric"] {
        background-color: #f8fafd;
        border-radius: 8px;
        padding: 10px;
        border: 1px solid #c5d5e6;
    }
    
    [data-testid="stMetric"] label {
        color: #1a5276 !important;
        font-weight: 600 !important;
    }
    
    [data-testid="stMetric"] .stMetricValue {
        color: #1a1a1a !important;
        font-weight: 700 !important;
    }
    
    .streamlit-expanderHeader {
        background-color: #f0f4f8 !important;
        color: #1a5276 !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
    }
    
    .streamlit-expanderContent {
        background-color: white !important;
        border-radius: 0 0 8px 8px !important;
    }
    
    hr {
        border-color: #c5d5e6 !important;
    }
</style>
""", unsafe_allow_html=True)

# -------------------------
# USUÁRIOS VIA TXT
# -------------------------
def carregar_usuarios(caminho="usuarios.txt"):
    if not os.path.exists(caminho):
        st.error("Arquivo usuarios.txt não encontrado.")
        st.stop()

    usuarios = {}

    with open(caminho, "r", encoding="utf-8") as f:
        for linha in f:
            linha = linha.strip()
            if not linha or ";" not in linha:
                continue
            usuario, senha = linha.split(";", 1)
            usuarios[usuario.strip()] = senha.strip()

    return usuarios


def tela_login():
    st.title("📐 Proanalisis v1.3")
    st.caption("Sistema de análise urbanística e geração de parecer técnico")

    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("### Acesso ao sistema")
        usuarios = carregar_usuarios()

        user = st.text_input("Usuário", key="login_user")
        senha = st.text_input("Senha", type="password", key="login_senha")

        if st.button("Entrar", use_container_width=True, key="btn_login"):
            if user in usuarios and usuarios[user] == senha:
                st.session_state["logado"] = True
                st.session_state["usuario"] = user
                st.rerun()
            else:
                st.error("Usuário ou senha inválidos.")


if "logado" not in st.session_state:
    st.session_state["logado"] = False

if not st.session_state["logado"]:
    tela_login()
    st.stop()

# -------------------------
# SIDEBAR
# -------------------------
st.sidebar.title("📐 Proanalisis v1.3")
st.sidebar.write(f"👤 {st.session_state['usuario']}")

if st.sidebar.button("🚪 Sair", use_container_width=True, key="btn_sair"):
    st.session_state["logado"] = False
    st.session_state.pop("dados_antigos", None)
    st.rerun()

# -------------------------
# HISTÓRICO
# -------------------------
def get_pasta_protocolo(protocolo):
    protocolo_limpo = protocolo.replace("/", "-").strip()
    return os.path.join("dados", protocolo_limpo)

def listar_analises(protocolo):
    pasta = get_pasta_protocolo(protocolo)
    if not os.path.exists(pasta):
        return []

    arquivos = os.listdir(pasta)
    analises = [f for f in arquivos if f.startswith("AN") and f.endswith(".json")]

    def ordem(nome):
        try:
            return int(nome.replace("AN", "").replace(".json", ""))
        except ValueError:
            return 999999

    analises.sort(key=ordem)
    return analises

def carregar_ultima_analise(protocolo):
    analises = listar_analises(protocolo)
    if not analises:
        return None

    caminho = os.path.join(get_pasta_protocolo(protocolo), analises[-1])
    with open(caminho, "r", encoding="utf-8") as f:
        return json.load(f)

def sugerir_proxima_analise(protocolo):
    analises = listar_analises(protocolo)
    if not analises:
        return "0"

    ultima = analises[-1]
    try:
        numero = int(ultima.replace("AN", "").replace(".json", ""))
        return str(numero + 1)
    except ValueError:
        return "0"

def salvar_historico(dados, respostas, observacoes, conclusao, analista, n_analise, arquivo_docx, pendencias_manuais):
    pasta = get_pasta_protocolo(dados["protocolo"])
    os.makedirs(pasta, exist_ok=True)

    base = f"AN{n_analise}"

    registro = {
        "protocolo": dados["protocolo"],
        "n_analise": n_analise,
        "data": datetime.now().strftime("%d/%m/%Y"),
        "analista": analista,
        "usuario": st.session_state["usuario"],
        "dados": dados,
        "respostas": respostas,
        "observacoes": observacoes,
        "pendencias_manuais": pendencias_manuais,
        "conclusao": conclusao
    }

    with open(os.path.join(pasta, f"{base}.json"), "w", encoding="utf-8") as f:
        json.dump(registro, f, indent=4, ensure_ascii=False)

    with open(os.path.join(pasta, f"{base}.docx"), "wb") as f:
        f.write(arquivo_docx.getvalue())

# -------------------------
# PERGUNTAS
# -------------------------
def carregar_perguntas_txt(caminho):
    if not os.path.exists(caminho):
        st.error("Arquivo perguntas.txt não encontrado.")
        st.stop()

    perguntas = []
    bloco = {}

    with open(caminho, "r", encoding="utf-8") as f:
        linhas = f.readlines()

    for linha in linhas:
        linha = linha.strip()

        if not linha:
            if bloco:
                perguntas.append(bloco)
                bloco = {}
            continue

        if linha.startswith("GRUPO:"):
            bloco["grupo"] = linha.replace("GRUPO:", "").strip()
        elif linha.startswith("ID:"):
            bloco["id"] = linha.replace("ID:", "").strip()
        elif linha.startswith("PERGUNTA:"):
            bloco["pergunta"] = linha.replace("PERGUNTA:", "").strip()
        elif linha.startswith("OPCOES:"):
            bloco["opcoes"] = [op.strip() for op in linha.replace("OPCOES:", "").strip().split(";")]
        elif linha.startswith("CONFORMES:"):
            bloco["conformes"] = [op.strip() for op in linha.replace("CONFORMES:", "").strip().split(";")]
        elif linha.startswith("REGRA_"):
            chave, valor = linha.split(":", 1)
            resposta = chave.replace("REGRA_", "").strip()
            bloco.setdefault("regras", {})[resposta] = {"texto": valor.strip()}

    if bloco:
        perguntas.append(bloco)

    return perguntas


def validar_ids_repetidos(perguntas):
    ids = [p.get("id", "").strip() for p in perguntas if p.get("id", "").strip()]
    return sorted({i for i in ids if ids.count(i) > 1})


perguntas = carregar_perguntas_txt("perguntas.txt")
ids_repetidos = validar_ids_repetidos(perguntas)
if ids_repetidos:
    st.error("Há IDs repetidos no perguntas.txt: " + ", ".join(ids_repetidos))
    st.stop()

# -------------------------
# FUNÇÕES
# -------------------------
def resposta_preenchida(valor):
    return valor not in ("", None, "Selecione...")

def definir_conclusao(respostas, pendencias_manuais=None):
    for p in perguntas:
        resposta = respostas.get(p["id"])
        if not resposta_preenchida(resposta):
            continue
        conformes = p.get("conformes", ["Sim", "Não se enquadra"])
        if resposta not in conformes and resposta in p.get("regras", {}):
            return "DESFAVORÁVEL"
    
    if pendencias_manuais:
        for grupo, pendencia in pendencias_manuais.items():
            if pendencia and pendencia.strip():
                return "DESFAVORÁVEL"
    
    return "FAVORÁVEL"


def montar_inconformidades_por_grupo(respostas, observacoes, pendencias_manuais=None):
    grupos = {}

    for p in perguntas:
        pid = p["id"]
        resp = respostas.get(pid)
        if not resposta_preenchida(resp):
            continue

        conformes = p.get("conformes", ["Sim", "Não se enquadra"])
        if resp not in conformes and resp in p.get("regras", {}):
            grupo = p["grupo"]
            texto = p["regras"][resp]["texto"]
            obs = observacoes.get(pid, "").strip()
            if obs:
                texto += f"\nObservação: {obs}"
            grupos.setdefault(grupo, []).append(texto)
    
    if pendencias_manuais:
        for grupo, pendencia in pendencias_manuais.items():
            if pendencia and pendencia.strip():
                grupos.setdefault(grupo, []).append(f"INCONFORMIDADE DIVERSA: {pendencia}")

    return grupos


def montar_inconformidades_rt(respostas, observacoes, pendencias_manuais=None):
    grupos = montar_inconformidades_por_grupo(respostas, observacoes, pendencias_manuais)

    rt = RichText()
    contador = 1

    if grupos:
        for grupo, itens in grupos.items():
            rt.add(grupo.upper(), bold=True)
            rt.add("\n\n")
            for item in itens:
                rt.add(f"{contador}. {item}")
                rt.add("\n\n")
                contador += 1
    else:
        rt.add("Não foram identificadas inconformidades.")

    return rt


def gerar_docx(dados, respostas, observacoes, conclusao, analista, matricula, setor, n_analise, pendencias_manuais=None):
    if not os.path.exists("modelo_parecer.docx"):
        st.error("Arquivo modelo_parecer.docx não encontrado.")
        st.stop()

    doc = DocxTemplate("modelo_parecer.docx")
    inconformidades_rt = montar_inconformidades_rt(respostas, observacoes, pendencias_manuais)

    matriculas_str = dados.get("matriculas", "")
    if isinstance(matriculas_str, list):
        matriculas_str = ", ".join(matriculas_str)

    context = {
        "protocolo": dados["protocolo"],
        "tipo": dados["tipo"],
        "interessado": dados["interessado"],
        "n_lotes": dados["n_lotes"],
        "matriculas": matriculas_str,
        "inconformidades": inconformidades_rt,
        "conclusao": conclusao,
        "data": f"Data: {datetime.now().strftime('%d/%m/%Y')}",
        "analista": f"Analista: {analista}",
        "matricula": matricula,
        "setor": setor,
        "n_analise": n_analise
    }

    doc.render(context)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def resumo_status_pergunta(p, resposta):
    if not resposta_preenchida(resposta):
        return "pendente"

    conformes = p.get("conformes", ["Sim", "Não se enquadra"])

    if resposta == "Não se enquadra":
        return "na"
    if resposta in conformes:
        return "conforme"
    if resposta in p.get("regras", {}):
        return "inconforme"
    return "neutro"


def progresso_percentual(respostas):
    total = len(perguntas)
    preenchidas = sum(1 for v in respostas.values() if resposta_preenchida(v))
    if total == 0:
        return 0, 0, 0.0
    pct = preenchidas / total
    return preenchidas, total, pct


def cor_progresso(pct):
    r = int(255 * (1 - pct))
    g = int(180 * pct + 60)
    b = 60
    return f"rgb({r},{g},{b})"


def render_progresso(preenchidas, total, pct, destino):
    cor = cor_progresso(pct)
    html = f"""
    <div><b>{preenchidas}/{total}</b> respostas preenchidas ({int(pct*100)}%)</div>
    <div class="progress-wrap">
        <div class="progress-bar" style="width:{pct*100:.1f}%; background:{cor};"></div>
    </div>
    """
    destino.markdown(html, unsafe_allow_html=True)

def inicializar_estados():
    if "dados_antigos" not in st.session_state:
        st.session_state["dados_antigos"] = None
    
    if "etapa" not in st.session_state:
        st.session_state["etapa"] = "1. Protocolo"
    
    if "protocolo" not in st.session_state:
        st.session_state["protocolo"] = ""
    if "tipo" not in st.session_state:
        st.session_state["tipo"] = "Loteamento"
    if "interessado" not in st.session_state:
        st.session_state["interessado"] = ""
    if "n_lotes" not in st.session_state:
        st.session_state["n_lotes"] = 1
    if "matriculas" not in st.session_state:
        st.session_state["matriculas"] = ""
    if "analista" not in st.session_state:
        st.session_state["analista"] = ""
    if "matricula_analista" not in st.session_state:
        st.session_state["matricula_analista"] = ""
    if "setor" not in st.session_state:
        st.session_state["setor"] = ""
    if "n_analise" not in st.session_state:
        st.session_state["n_analise"] = ""
    if "pendencias_manuais" not in st.session_state:
        st.session_state["pendencias_manuais"] = {}

# -------------------------
# INICIALIZAÇÃO
# -------------------------
inicializar_estados()

# -------------------------
# CABEÇALHO
# -------------------------
st.title("📐 Proanalisis v1.3")
st.caption("Análise urbanística padronizada com geração de parecer técnico")

etapas = [
    "1. Protocolo",
    "2. Analista",
    "3. Análise",
    "4. Revisão",
    "5. Gerar parecer"
]

etapa_atual = st.sidebar.radio("📋 Etapas", etapas, index=etapas.index(st.session_state["etapa"]))
if etapa_atual != st.session_state["etapa"]:
    st.session_state["etapa"] = etapa_atual
    st.rerun()

# -------------------------
# ETAPA 1
# -------------------------
if st.session_state["etapa"] == "1. Protocolo":
    st.header("📋 Dados do protocolo")

    c1, c2 = st.columns([2, 1])
    with c1:
        protocolo = st.text_input("N° Protocolo", value=st.session_state["protocolo"], key="protocolo_input")
        if protocolo != st.session_state["protocolo"]:
            st.session_state["protocolo"] = protocolo
    with c2:
        st.markdown("<div class='small-muted'>Use o mesmo protocolo para continuar uma análise já existente.</div>", unsafe_allow_html=True)

    st.subheader("🏢 Dados do empreendimento")
    
    tipo = st.selectbox(
        "Tipo do Empreendimento",
        ["Loteamento", "Condomínio fechado de lotes"],
        index=0 if st.session_state["tipo"] == "Loteamento" else 1,
        key="tipo_select"
    )
    st.session_state["tipo"] = tipo
    
    interessado = st.text_input("Requerente", value=st.session_state["interessado"], key="interessado_input")
    st.session_state["interessado"] = interessado
    
    n_lotes = st.number_input("Número de Lotes", min_value=1, value=st.session_state["n_lotes"], key="n_lotes_input")
    st.session_state["n_lotes"] = n_lotes
    
    matriculas = st.text_area(
        "Matrícula(s) do Empreendimento",
        value=st.session_state["matriculas"],
        key="matriculas_input",
        placeholder="Digite a(s) matrícula(s) separadas por vírgula. Ex: 12345, 67890",
        help="Informe a(s) matrícula(s) do imóvel no cartório de registro de imóveis"
    )
    st.session_state["matriculas"] = matriculas

    st.markdown("---")
    
    if st.session_state["protocolo"]:
        ultima = carregar_ultima_analise(st.session_state["protocolo"])
        if ultima:
            st.info(f"📋 Última análise encontrada: AN{ultima['n_analise']} - Data: {ultima.get('data', 'Data não disponível')}")
            col_a, col_b = st.columns([1, 1])
            with col_a:
                if st.button("▶️ Continuar análise", use_container_width=True):
                    st.session_state["dados_antigos"] = ultima
                    st.session_state["tipo"] = ultima["dados"].get("tipo", "Loteamento")
                    st.session_state["interessado"] = ultima["dados"].get("interessado", "")
                    st.session_state["n_lotes"] = int(ultima["dados"].get("n_lotes", 1))
                    st.session_state["matriculas"] = ultima["dados"].get("matriculas", "")
                    st.session_state["etapa"] = "2. Analista"
                    st.rerun()
            with col_b:
                if st.button("➕ Iniciar nova análise", use_container_width=True):
                    st.session_state["dados_antigos"] = None
                    st.session_state["etapa"] = "2. Analista"
                    st.rerun()
        else:
            st.success("✅ Nenhum histórico encontrado para este protocolo.")
            if st.button("Prosseguir →", use_container_width=True, type="primary"):
                if not st.session_state["protocolo"]:
                    st.error("⚠️ Por favor, informe o número do protocolo.")
                elif not st.session_state["interessado"]:
                    st.error("⚠️ Por favor, informe o requerente.")
                else:
                    st.session_state["dados_antigos"] = None
                    st.session_state["etapa"] = "2. Analista"
                    st.rerun()
    else:
        st.warning("⚠️ Informe o número do protocolo para continuar.")

# -------------------------
# ETAPA 2
# -------------------------
elif st.session_state["etapa"] == "2. Analista":
    st.header("👤 Dados do analista")
    
    st.info(f"📌 Protocolo: **{st.session_state['protocolo']}** | Empreendimento: **{st.session_state['interessado']}**")

    c1, c2 = st.columns(2)
    with c1:
        analista = st.text_input("Nome do Analista", value=st.session_state["analista"], key="analista_input")
        st.session_state["analista"] = analista
        
        matricula_analista = st.text_input("Matrícula do Analista", value=st.session_state["matricula_analista"], key="matricula_analista_input")
        st.session_state["matricula_analista"] = matricula_analista
        
    with c2:
        setor = st.text_input("Setor", value=st.session_state["setor"], key="setor_input")
        st.session_state["setor"] = setor
        
        n_analise_sugerida = sugerir_proxima_analise(st.session_state["protocolo"]) if st.session_state["protocolo"] else "1"
        if not st.session_state["n_analise"]:
            st.session_state["n_analise"] = n_analise_sugerida
        
        n_analise = st.text_input("Nº da Análise", value=st.session_state["n_analise"], key="n_analise_input")
        st.session_state["n_analise"] = n_analise

    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("← Voltar", use_container_width=True):
            st.session_state["etapa"] = "1. Protocolo"
            st.rerun()
    with col2:
        if st.button("Prosseguir →", use_container_width=True, type="primary"):
            if not st.session_state["analista"]:
                st.error("⚠️ Por favor, informe o nome do analista.")
            elif not st.session_state["n_analise"]:
                st.error("⚠️ Por favor, informe o número da análise.")
            else:
                st.session_state["etapa"] = "3. Análise"
                st.rerun()

# -------------------------
# ETAPA 3
# -------------------------
elif st.session_state["etapa"] == "3. Análise":
    st.header("🔍 Análise técnica")
    
    st.info(f"📌 Protocolo: **{st.session_state['protocolo']}** | Analista: **{st.session_state['analista']}** | Análise Nº: **{st.session_state['n_analise']}**")

    respostas = {}
    observacoes = {}
    grupos_ui = {}
    pendencias_manuais = st.session_state.get("pendencias_manuais", {})

    for idx, p in enumerate(perguntas):
        grupos_ui.setdefault(p["grupo"], []).append((idx, p))

    inconformes_sidebar = []

    for grupo, lista in grupos_ui.items():
        with st.expander(f"📁 {grupo}", expanded=False):
            for idx, p in lista:
                pid = p["id"]
                valor_padrao = None
                obs_padrao = ""

                if st.session_state["dados_antigos"]:
                    valor_padrao = st.session_state["dados_antigos"]["respostas"].get(pid)
                    obs_padrao = st.session_state["dados_antigos"]["observacoes"].get(pid, "")

                chave_base = f"{idx}_{grupo}_{pid}"
                chave_resp = f"resp_{chave_base}"
                chave_obs = f"obs_{chave_base}"

                opcoes_exibicao = ["Selecione..."] + p["opcoes"]

                if chave_resp not in st.session_state:
                    st.session_state[chave_resp] = valor_padrao if valor_padrao in p["opcoes"] else "Selecione..."
                if chave_obs not in st.session_state:
                    st.session_state[chave_obs] = obs_padrao

                indice_padrao = opcoes_exibicao.index(st.session_state[chave_resp]) if st.session_state[chave_resp] in opcoes_exibicao else 0

                col_pergunta, col_status = st.columns([3, 1])
                
                with col_pergunta:
                    resposta = st.selectbox(
                        p["pergunta"],
                        opcoes_exibicao,
                        index=indice_padrao,
                        key=chave_resp,
                        help=f"ID: {pid}"
                    )
                    respostas[pid] = resposta
                
                obs = st.text_area(
                    "📝 Observação (opcional)",
                    key=chave_obs,
                    height=68,
                    placeholder="Registre detalhes adicionais sobre esta resposta..."
                )
                observacoes[pid] = obs

                status = resumo_status_pergunta(p, respostas[pid])
                
                with col_status:
                    if status == "inconforme":
                        st.markdown("""
                        <div class='status-badge' style='background-color:#fdeaea; border-left: 4px solid #b42318;'>
                            <span class='status-icon'>⛔</span>
                            <div class='status-text' style='color: #b42318;'>INCONFORME</div>
                        </div>
                        """, unsafe_allow_html=True)
                        inconformes_sidebar.append(p["pergunta"])
                    elif status == "na":
                        st.markdown("""
                        <div class='status-badge' style='background-color:#eef4ff; border-left: 4px solid #1a5276;'>
                            <span class='status-icon'>ℹ️</span>
                            <div class='status-text' style='color: #1a5276;'>NÃO SE ENQUADRA</div>
                        </div>
                        """, unsafe_allow_html=True)
                    elif status == "conforme":
                        st.markdown("""
                        <div class='status-badge' style='background-color:#ecfdf3; border-left: 4px solid #067647;'>
                            <span class='status-icon'>✅</span>
                            <div class='status-text' style='color: #067647;'>CONFORME</div>
                        </div>
                        """, unsafe_allow_html=True)
                    else:
                        st.markdown("""
                        <div class='status-badge' style='background-color:#fffaeb; border-left: 4px solid #b54708;'>
                            <span class='status-icon'>⏳</span>
                            <div class='status-text' style='color: #b54708;'>PENDENTE</div>
                        </div>
                        """, unsafe_allow_html=True)
                
                st.markdown("---")
            
            st.markdown("### 📝 Inconformidades Diversas")
            st.caption("Registre aqui quaisquer inconformidades adicionais não cobertas pelas perguntas acima")
            
            chave_pendencia = f"pendencia_{grupo}"
            valor_padrao_pendencia = pendencias_manuais.get(grupo, "")
            
            pendencia = st.text_area(
                f"Inconformidades diversas para o grupo: {grupo}",
                value=valor_padrao_pendencia,
                key=chave_pendencia,
                height=80,
                placeholder="Ex: Documentação incompleta, falta de assinatura, necessidade de complementação de informações..."
            )
            pendencias_manuais[grupo] = pendencia
            
            if pendencia and pendencia.strip():
                inconformes_sidebar.append(f"{grupo} - Inconformidade Diversa")
            
            st.markdown("---")

    st.session_state["pendencias_manuais"] = pendencias_manuais

    preenchidas, total, pct = progresso_percentual(respostas)

    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📊 Progresso")
    render_progresso(preenchidas, total, pct, st.sidebar)

    st.sidebar.markdown("### ⚠️ Inconformidades")
    if inconformes_sidebar:
        for item in inconformes_sidebar[:20]:
            st.sidebar.write(f"- {item}")
    else:
        st.sidebar.write("Nenhuma até o momento.")

    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("← Voltar", use_container_width=True):
            st.session_state["etapa"] = "2. Analista"
            st.rerun()
    with col2:
        if st.button("Prosseguir para revisão →", use_container_width=True, type="primary"):
            st.session_state["respostas_temp"] = respostas
            st.session_state["observacoes_temp"] = observacoes
            st.session_state["etapa"] = "4. Revisão"
            st.rerun()

# -------------------------
# ETAPA 4
# -------------------------
elif st.session_state["etapa"] == "4. Revisão":
    st.header("📋 Revisão da análise")
    
    if "respostas_temp" in st.session_state:
        respostas = st.session_state["respostas_temp"]
        observacoes = st.session_state["observacoes_temp"]
    else:
        respostas = {}
        observacoes = {}
        for idx, p in enumerate(perguntas):
            chave_base = f"{idx}_{p['grupo']}_{p['id']}"
            respostas[p["id"]] = st.session_state.get(f"resp_{chave_base}", "Selecione...")
            observacoes[p["id"]] = st.session_state.get(f"obs_{chave_base}", "")

    pendencias_manuais = st.session_state.get("pendencias_manuais", {})

    preenchidas, total, pct = progresso_percentual(respostas)
    render_progresso(preenchidas, total, pct, st)

    conclusao = definir_conclusao(respostas, pendencias_manuais)
    grupos_inconformes = montar_inconformidades_por_grupo(respostas, observacoes, pendencias_manuais)

    col1, col2 = st.columns([1.4, 1])

    with col1:
        st.subheader("📌 Resumo geral")
        st.write(f"**Protocolo:** {st.session_state.get('protocolo', '')}")
        st.write(f"**Requerente:** {st.session_state.get('interessado', '')}")
        st.write(f"**Tipo:** {st.session_state.get('tipo', '')}")
        st.write(f"**Matrícula(s):** {st.session_state.get('matriculas', '')}")
        st.write(f"**Analista:** {st.session_state.get('analista', '')}")
        st.write(f"**Nº da análise:** {st.session_state.get('n_analise', '')}")
        
        if conclusao == "FAVORÁVEL":
            st.success(f"✅ **Conclusão preliminar:** {conclusao}")
        else:
            st.error(f"❌ **Conclusão preliminar:** {conclusao}")

    with col2:
        st.subheader("📊 Contagem")
        total_inconformes = sum(len(v) for v in grupos_inconformes.values())
        st.metric("Perguntas", len(perguntas))
        st.metric("Respondidas", preenchidas)
        st.metric("Inconformidades", total_inconformes)

    st.subheader("⚠️ Inconformidades identificadas")
    if grupos_inconformes:
        for grupo, itens in grupos_inconformes.items():
            st.markdown(f"#### {grupo}")
            for i, item in enumerate(itens, start=1):
                st.markdown(f"""
                <div style="background-color: #f8fafd; border: 1px solid #c5d5e6; border-radius: 10px; padding: 12px 16px; margin: 8px 0; color: #1a1a1a;">
                    <b style="color: #1a5276;">{i}.</b> 
                    <span style="color: #1a1a1a;">{item.replace(chr(10), '<br>')}</span>
                </div>
                """, unsafe_allow_html=True)
    else:
        st.success("✅ Não foram identificadas inconformidades.")

    if st.session_state.get("dados_antigos"):
        st.subheader("🔄 Alterações em relação à análise anterior")
        houve_alteracao = False
        dados_antigos = st.session_state["dados_antigos"]

        for p in perguntas:
            pid = p["id"]
            antiga = dados_antigos["respostas"].get(pid)
            atual = respostas.get(pid)
            obs_antiga = dados_antigos["observacoes"].get(pid, "").strip()
            obs_atual = observacoes.get(pid, "").strip()

            if antiga != atual or obs_antiga != obs_atual:
                houve_alteracao = True
                st.warning(f"**{pid}:** resposta '{antiga}' → '{atual}' | observação '{obs_antiga}' → '{obs_atual}'")

        if not houve_alteracao:
            st.success("✅ Nenhuma alteração identificada em relação à análise carregada.")

    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("← Voltar para análise", use_container_width=True):
            st.session_state["etapa"] = "3. Análise"
            st.rerun()
    with col2:
        if st.button("Prosseguir para geração →", use_container_width=True, type="primary"):
            st.session_state["etapa"] = "5. Gerar parecer"
            st.rerun()

# -------------------------
# ETAPA 5
# -------------------------
elif st.session_state["etapa"] == "5. Gerar parecer":
    st.header("📄 Geração do parecer")
    
    if "respostas_temp" in st.session_state:
        respostas = st.session_state["respostas_temp"]
        observacoes = st.session_state["observacoes_temp"]
    else:
        respostas = {}
        observacoes = {}
        for idx, p in enumerate(perguntas):
            chave_base = f"{idx}_{p['grupo']}_{p['id']}"
            respostas[p["id"]] = st.session_state.get(f"resp_{chave_base}", "Selecione...")
            observacoes[p["id"]] = st.session_state.get(f"obs_{chave_base}", "")

    pendencias_manuais = st.session_state.get("pendencias_manuais", {})

    preenchidas, total, pct = progresso_percentual(respostas)
    render_progresso(preenchidas, total, pct, st)

    dados = {
        "protocolo": st.session_state.get("protocolo", ""),
        "tipo": st.session_state.get("tipo", ""),
        "interessado": st.session_state.get("interessado", ""),
        "n_lotes": st.session_state.get("n_lotes", 1),
        "matriculas": st.session_state.get("matriculas", "")
    }

    conclusao = definir_conclusao(respostas, pendencias_manuais)

    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**📌 Protocolo:** {dados['protocolo']}")
        st.write(f"**👤 Requerente:** {dados['interessado']}")
        st.write(f"**🏢 Tipo:** {dados['tipo']}")
        st.write(f"**📋 Matrícula(s):** {dados['matriculas']}")
    with col2:
        st.write(f"**🔢 Nº Lotes:** {dados['n_lotes']}")
        st.write(f"**👨‍💼 Analista:** {st.session_state.get('analista', '')}")
        st.write(f"**🔢 Matrícula Analista:** {st.session_state.get('matricula_analista', '')}")
        st.write(f"**🔍 Nº Análise:** {st.session_state.get('n_analise', '')}")
        
    if conclusao == "FAVORÁVEL":
        st.success(f"✅ **Conclusão final:** {conclusao}")
    else:
        st.error(f"❌ **Conclusão final:** {conclusao}")

    campos_invalidos = []
    if not dados["protocolo"]:
        campos_invalidos.append("Protocolo")
    if not st.session_state.get("analista"):
        campos_invalidos.append("Analista")
    if not st.session_state.get("n_analise"):
        campos_invalidos.append("Número da Análise")
    
    if campos_invalidos:
        st.error(f"⚠️ Campos obrigatórios pendentes: {', '.join(campos_invalidos)}")
    
    if preenchidas < total:
        st.warning(f"⚠️ Atenção: {total - preenchidas} perguntas ainda estão pendentes. Revise antes de gerar o parecer.")
        st.info("💡 Você pode voltar para a etapa de Análise para responder as perguntas pendentes.")
    
    pendencias_registradas = {k: v for k, v in pendencias_manuais.items() if v and v.strip()}
    if pendencias_registradas:
        st.warning("⚠️ Inconformidades diversas registradas:")
        for grupo, pendencia in pendencias_registradas.items():
            st.markdown(f"- **{grupo}:** {pendencia}")
    
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("← Voltar para revisão", use_container_width=True):
            st.session_state["etapa"] = "4. Revisão"
            st.rerun()
    with col2:
        if st.button("📄 Gerar Parecer Técnico", use_container_width=True, type="primary"):
            if not dados["protocolo"]:
                st.error("❌ Protocolo não informado. Volte à etapa 1 e informe o número do protocolo.")
                st.stop()
            
            if not st.session_state.get("analista"):
                st.error("❌ Nome do analista não informado. Volte à etapa 2 e preencha seus dados.")
                st.stop()
            
            if not st.session_state.get("n_analise"):
                st.error("❌ Número da análise não informado. Volte à etapa 2 e preencha o número da análise.")
                st.stop()
            
            if not os.path.exists("modelo_parecer.docx"):
                st.error("❌ Arquivo 'modelo_parecer.docx' não encontrado. Verifique se o arquivo está no diretório correto.")
                st.stop()
            
            try:
                with st.spinner("Gerando parecer técnico... Aguarde."):
                    arquivo = gerar_docx(
                        dados=dados,
                        respostas=respostas,
                        observacoes=observacoes,
                        conclusao=conclusao,
                        analista=st.session_state.get("analista", ""),
                        matricula=st.session_state.get("matricula_analista", ""),
                        setor=st.session_state.get("setor", ""),
                        n_analise=st.session_state.get("n_analise", ""),
                        pendencias_manuais=pendencias_manuais
                    )

                    salvar_historico(
                        dados=dados,
                        respostas=respostas,
                        observacoes=observacoes,
                        conclusao=conclusao,
                        analista=st.session_state.get("analista", ""),
                        n_analise=st.session_state.get("n_analise", ""),
                        arquivo_docx=arquivo,
                        pendencias_manuais=pendencias_manuais
                    )

                    protocolo_limpo = dados["protocolo"].replace("/", "-")
                    data_arquivo = datetime.now().strftime("%d-%m-%Y")
                    analise_str = f"AN{st.session_state.get('n_analise', '0')}"
                    nome_arquivo = f"PU_{protocolo_limpo}_{data_arquivo}_{analise_str}.docx"

                    st.success("✅ Parecer gerado e histórico salvo com sucesso!")
                    
                    st.download_button(
                        label="⬇️ Baixar parecer (.docx)",
                        data=arquivo,
                        file_name=nome_arquivo,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                    
            except Exception as e:
                st.error(f"❌ Erro ao gerar o parecer: {str(e)}")
                st.error("Verifique se:")
                st.error("1. O arquivo 'modelo_parecer.docx' existe no diretório")
                st.error("2. O arquivo 'modelo_parecer.docx' não está corrompido")
                st.error("3. Todas as variáveis do template estão corretas")
                st.exception(e)
