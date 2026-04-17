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

st.markdown("""
<style>
.block-container {
    padding-top: 1.2rem;
    padding-bottom: 1.2rem;
}
.stButton > button {
    border-radius: 8px;
    font-weight: 600;
}
.small-muted {
    color: #666;
    font-size: 0.9rem;
}
.card {
    padding: 0.8rem 1rem;
    border: 1px solid #e6e6e6;
    border-radius: 10px;
    background: #fafafa;
    margin-bottom: 0.6rem;
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
    st.title("Proanalisis v1.3")
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
st.sidebar.title("Proanalisis v1.3")
st.sidebar.write(f"👤 {st.session_state['usuario']}")

if st.sidebar.button("Sair", use_container_width=True, key="btn_sair"):
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

def salvar_historico(dados, respostas, observacoes, conclusao, analista, n_analise, arquivo_docx):
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

def definir_conclusao(respostas):
    for p in perguntas:
        resposta = respostas.get(p["id"])
        if not resposta_preenchida(resposta):
            continue
        conformes = p.get("conformes", ["Sim", "Não se enquadra"])
        if resposta not in conformes and resposta in p.get("regras", {}):
            return "DESFAVORÁVEL"
    return "FAVORÁVEL"


def montar_inconformidades_por_grupo(respostas, observacoes):
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

    return grupos


def montar_inconformidades_rt(respostas, observacoes):
    grupos = montar_inconformidades_por_grupo(respostas, observacoes)

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


def gerar_docx(dados, respostas, observacoes, conclusao, analista, matricula, setor, n_analise):
    if not os.path.exists("modelo_parecer.docx"):
        st.error("Arquivo modelo_parecer.docx não encontrado.")
        st.stop()

    doc = DocxTemplate("modelo_parecer.docx")
    inconformidades_rt = montar_inconformidades_rt(respostas, observacoes)

    context = {
        "protocolo": dados["protocolo"],
        "tipo": dados["tipo"],
        "interessado": dados["interessado"],
        "n_lotes": dados["n_lotes"],
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
    # 0 = vermelho, 1 = verde
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

# -------------------------
# ESTADO INICIAL
# -------------------------
if "dados_antigos" not in st.session_state:
    st.session_state["dados_antigos"] = None

if "etapa" not in st.session_state:
    st.session_state["etapa"] = "1. Protocolo"

# -------------------------
# CABEÇALHO
# -------------------------
st.title("Proanalisis v1.3")
st.caption("Análise urbanística padronizada com geração de parecer técnico")

etapas = [
    "1. Protocolo",
    "2. Analista",
    "3. Análise",
    "4. Revisão",
    "5. Gerar parecer"
]

st.session_state["etapa"] = st.sidebar.radio("Etapas", etapas, index=etapas.index(st.session_state["etapa"]))

# -------------------------
# CAMPOS BASE
# -------------------------
dados_antigos = st.session_state.get("dados_antigos")

if "protocolo" not in st.session_state:
    st.session_state["protocolo"] = ""
if "tipo" not in st.session_state:
    st.session_state["tipo"] = "Loteamento"
if "interessado" not in st.session_state:
    st.session_state["interessado"] = ""
if "n_lotes" not in st.session_state:
    st.session_state["n_lotes"] = 1
if "analista" not in st.session_state:
    st.session_state["analista"] = ""
if "matricula" not in st.session_state:
    st.session_state["matricula"] = ""
if "setor" not in st.session_state:
    st.session_state["setor"] = ""

# -------------------------
# ETAPA 1
# -------------------------
if st.session_state["etapa"] == "1. Protocolo":
    st.header("Dados do protocolo")

    c1, c2 = st.columns([2, 1])
    with c1:
        protocolo = st.text_input("N° Protocolo", key="protocolo")
    with c2:
        st.markdown("<div class='small-muted'>Use o mesmo protocolo para continuar uma análise já existente.</div>", unsafe_allow_html=True)

    if protocolo:
        ultima = carregar_ultima_analise(protocolo)
        if ultima:
            st.info(f"Última análise encontrada: AN{ultima['n_analise']}")
            col_a, col_b = st.columns([1, 1])
            with col_a:
                if st.button("▶️ Continuar análise", use_container_width=True):
                    st.session_state["dados_antigos"] = ultima
                    st.session_state["tipo"] = ultima["dados"].get("tipo", "Loteamento")
                    st.session_state["interessado"] = ultima["dados"].get("interessado", "")
                    st.session_state["n_lotes"] = int(ultima["dados"].get("n_lotes", 1))
                    st.session_state["etapa"] = "2. Analista"
                    st.rerun()
            with col_b:
                if st.button("➕ Iniciar nova análise", use_container_width=True):
                    st.session_state["dados_antigos"] = None
                    st.session_state["etapa"] = "2. Analista"
                    st.rerun()
        else:
            st.success("Nenhum histórico encontrado para este protocolo.")
            if st.button("Prosseguir", use_container_width=True):
                st.session_state["dados_antigos"] = None
                st.session_state["etapa"] = "2. Analista"
                st.rerun()

    st.subheader("Dados do empreendimento")
    st.selectbox(
        "Tipo do Empreendimento",
        ["Loteamento", "Condomínio fechado de lotes"],
        key="tipo"
    )
    st.text_input("Requerente", key="interessado")
    st.number_input("Número de Lotes", min_value=1, key="n_lotes")

# -------------------------
# ETAPA 2
# -------------------------
elif st.session_state["etapa"] == "2. Analista":
    st.header("Dados do analista")

    c1, c2 = st.columns(2)
    with c1:
        st.text_input("Nome do Analista", key="analista")
        st.text_input("Matrícula", key="matricula")
    with c2:
        st.text_input("Setor", key="setor")
        n_analise_sugerida = sugerir_proxima_analise(st.session_state["protocolo"]) if st.session_state["protocolo"] else "0"
        if "n_analise" not in st.session_state:
            st.session_state["n_analise"] = n_analise_sugerida
        st.text_input("Nº da Análise", key="n_analise")

    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("← Voltar", use_container_width=True):
            st.session_state["etapa"] = "1. Protocolo"
            st.rerun()
    with col2:
        if st.button("Prosseguir →", use_container_width=True):
            st.session_state["etapa"] = "3. Análise"
            st.rerun()

# -------------------------
# ETAPA 3
# -------------------------
elif st.session_state["etapa"] == "3. Análise":
    st.header("Análise técnica")

    respostas = {}
    observacoes = {}
    grupos_ui = {}

    for idx, p in enumerate(perguntas):
        grupos_ui.setdefault(p["grupo"], []).append((idx, p))

    inconformes_sidebar = []

    for grupo, lista in grupos_ui.items():
        with st.expander(grupo, expanded=False):
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

                col1, col2 = st.columns([1.2, 1])
                with col1:
                    respostas[pid] = st.selectbox(
                        p["pergunta"],
                        opcoes_exibicao,
                        index=indice_padrao,
                        key=chave_resp,
                        help=f"ID: {pid}"
                    )
                with col2:
                    observacoes[pid] = st.text_area(
                        "Observação",
                        key=chave_obs,
                        height=80
                    )

                status = resumo_status_pergunta(p, respostas[pid])
                if status == "inconforme":
                    st.error("Inconformidade identificada")
                    inconformes_sidebar.append(p["pergunta"])
                elif status == "na":
                    st.info("Não se enquadra")
                elif status == "conforme":
                    st.success("Conforme")
                else:
                    st.warning("Pendente de resposta")

    preenchidas, total, pct = progresso_percentual(respostas)

    st.sidebar.markdown("---")
    st.sidebar.markdown("### Progresso")
    render_progresso(preenchidas, total, pct, st.sidebar)

    st.sidebar.markdown("### ⚠ Inconformidades")
    if inconformes_sidebar:
        for item in inconformes_sidebar[:20]:
            st.sidebar.write(f"- {item}")
    else:
        st.sidebar.write("Nenhuma até o momento.")

    if st.button("Prosseguir para revisão →", use_container_width=True):
        st.session_state["etapa"] = "4. Revisão"
        st.rerun()

# -------------------------
# ETAPA 4
# -------------------------
elif st.session_state["etapa"] == "4. Revisão":
    st.header("Revisão da análise")

    respostas = {}
    observacoes = {}
    for idx, p in enumerate(perguntas):
        chave_base = f"{idx}_{p['grupo']}_{p['id']}"
        respostas[p["id"]] = st.session_state.get(f"resp_{chave_base}", "Selecione...")
        observacoes[p["id"]] = st.session_state.get(f"obs_{chave_base}", "")

    preenchidas, total, pct = progresso_percentual(respostas)
    render_progresso(preenchidas, total, pct, st)

    conclusao = definir_conclusao(respostas)
    grupos_inconformes = montar_inconformidades_por_grupo(respostas, observacoes)

    col1, col2 = st.columns([1.4, 1])

    with col1:
        st.subheader("Resumo geral")
        st.write(f"**Protocolo:** {st.session_state.get('protocolo', '')}")
        st.write(f"**Requerente:** {st.session_state.get('interessado', '')}")
        st.write(f"**Tipo:** {st.session_state.get('tipo', '')}")
        st.write(f"**Nº da análise:** {st.session_state.get('n_analise', '')}")
        st.write(f"**Conclusão preliminar:** {conclusao}")

    with col2:
        st.subheader("Contagem")
        total_inconformes = sum(len(v) for v in grupos_inconformes.values())
        st.metric("Perguntas", len(perguntas))
        st.metric("Respondidas", preenchidas)
        st.metric("Inconformidades", total_inconformes)

    st.subheader("Inconformidades identificadas")
    if grupos_inconformes:
        for grupo, itens in grupos_inconformes.items():
            st.markdown(f"#### {grupo}")
            for i, item in enumerate(itens, start=1):
                st.markdown(f"<div class='card'><b>{i}.</b> {item.replace(chr(10), '<br>')}</div>", unsafe_allow_html=True)
    else:
        st.success("Não foram identificadas inconformidades.")

    if st.session_state.get("dados_antigos"):
        st.subheader("Alterações em relação à análise anterior")
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
                st.warning(
                    f"{pid}: resposta '{antiga}' → '{atual}' | observação '{obs_antiga}' → '{obs_atual}'"
                )

        if not houve_alteracao:
            st.success("Nenhuma alteração identificada em relação à análise carregada.")

    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("← Voltar para análise", use_container_width=True):
            st.session_state["etapa"] = "3. Análise"
            st.rerun()
    with col2:
        if st.button("Prosseguir para geração →", use_container_width=True):
            st.session_state["etapa"] = "5. Gerar parecer"
            st.rerun()

# -------------------------
# ETAPA 5
# -------------------------
elif st.session_state["etapa"] == "5. Gerar parecer":
    st.header("Geração do parecer")

    respostas = {}
    observacoes = {}
    for idx, p in enumerate(perguntas):
        chave_base = f"{idx}_{p['grupo']}_{p['id']}"
        respostas[p["id"]] = st.session_state.get(f"resp_{chave_base}", "Selecione...")
        observacoes[p["id"]] = st.session_state.get(f"obs_{chave_base}", "")

    preenchidas, total, pct = progresso_percentual(respostas)
    render_progresso(preenchidas, total, pct, st)

    dados = {
        "protocolo": st.session_state.get("protocolo", ""),
        "tipo": st.session_state.get("tipo", ""),
        "interessado": st.session_state.get("interessado", ""),
        "n_lotes": st.session_state.get("n_lotes", 1)
    }

    conclusao = definir_conclusao(respostas)

    st.write(f"**Protocolo:** {dados['protocolo']}")
    st.write(f"**Conclusão final:** {conclusao}")

    if st.button("📄 Gerar Parecer Técnico", use_container_width=True):
        if not str(dados["protocolo"]).strip():
            st.error("Informe o número do protocolo.")
            st.stop()

        if not str(st.session_state.get("n_analise", "")).strip():
            st.error("Informe o número da análise.")
            st.stop()

        arquivo = gerar_docx(
            dados=dados,
            respostas=respostas,
            observacoes=observacoes,
            conclusao=conclusao,
            analista=st.session_state.get("analista", ""),
            matricula=st.session_state.get("matricula", ""),
            setor=st.session_state.get("setor", ""),
            n_analise=st.session_state.get("n_analise", "")
        )

        salvar_historico(
            dados=dados,
            respostas=respostas,
            observacoes=observacoes,
            conclusao=conclusao,
            analista=st.session_state.get("analista", ""),
            n_analise=st.session_state.get("n_analise", ""),
            arquivo_docx=arquivo
        )

        protocolo_limpo = dados["protocolo"].replace("/", "-")
        data_arquivo = datetime.now().strftime("%d-%m-%Y")
        analise_str = f"AN{st.session_state.get('n_analise', '0')}"
        nome_arquivo = f"PU_{protocolo_limpo}_{data_arquivo}_{analise_str}.docx"

        st.success("Parecer gerado e histórico salvo com sucesso.")
        st.download_button(
            label="⬇️ Baixar parecer (.docx)",
            data=arquivo,
            file_name=nome_arquivo,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

    if st.button("← Voltar para revisão", use_container_width=True):
        st.session_state["etapa"] = "4. Revisão"
        st.rerun()