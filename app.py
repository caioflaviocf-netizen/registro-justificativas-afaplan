import os
import json
import calendar
from datetime import datetime, date, timedelta
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from filelock import FileLock

# ==============================================================================
# 1. CONFIGURAÇÃO DA PÁGINA E TEMA CLARO (FIEL AO MODELO DE REFERÊNCIA)
# ==============================================================================
st.set_page_config(
    page_title="Registro de Justificativas de Desempenho - Afaplan",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    /* Tema claro uniforme em qualquer dispositivo */
    [data-testid="stAppViewContainer"],
    [data-testid="stHeader"],
    .stApp {
        background-color: #f4f6f9 !important;
        color: #1f2933 !important;
    }

    /* Garante que a barra superior do Streamlit não corte o cabeçalho */
    [data-testid="stHeader"] {
        background-color: rgba(244, 246, 249, 0.85) !important;
        height: 3.2rem !important;
    }

    .block-container {
        padding-top: 3.4rem !important;
        padding-bottom: 3rem !important;
        padding-left: 1.5rem !important;
        padding-right: 1.5rem !important;
        max-width: 1180px !important;
        margin: 0 auto !important;
    }

    /* Cabeçalho superior fiel ao modelo de referência */
    .afaplan-banner {
        display: flex;
        align-items: stretch;
        background-color: #ffffff;
        box-shadow: 0 2px 6px rgba(0,0,0,0.08);
        border-radius: 6px;
        overflow: hidden;
        min-height: 76px;
        width: 100%;
        margin-bottom: 14px;
        border: 1px solid #d9e2ec;
    }
    .afaplan-logo-box {
        background-color: #ffffff;
        padding: 12px 26px;
        display: flex;
        align-items: center;
        justify-content: center;
        min-width: 215px;
    }
    .afaplan-title-box {
        flex: 1;
        background-color: #1560bd;
        color: #ffffff;
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 12px 26px;
        gap: 12px;
        flex-wrap: wrap;
    }
    .afaplan-title-box h1 {
        color: #ffffff !important;
        font-size: 1.35rem !important;
        font-weight: 700 !important;
        margin: 0 !important;
        padding: 0 !important;
        letter-spacing: 0.5px;
        line-height: 1.2 !important;
    }
    .afaplan-badge {
        background: rgba(255,255,255,0.18);
        color: #ffffff;
        padding: 5px 14px;
        border-radius: 20px;
        font-size: 0.84rem;
        font-weight: 600;
    }

    /* Rótulos e inputs */
    label, .stSelectbox label, .stTextInput label, .stMarkdown p {
        color: #1f2933 !important;
    }
    div[data-baseweb="select"] > div,
    div[data-baseweb="input"] > div,
    input[type="text"],
    input[type="password"] {
        background-color: #ffffff !important;
        color: #102a43 !important;
        border-color: #cbd2d9 !important;
    }

    /* Card de cada linha diária (Data/Hora + Caixa de texto centralizada/alinhada) */
    .dia-card-box {
        background-color: #ffffff;
        border: 1px solid #e4e7eb;
        border-radius: 6px;
        padding: 10px 14px;
        margin-bottom: 8px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }
    .dia-card-box.feriado-destaque {
        background-color: #fffdf5;
        border: 1px solid #f59f00;
        border-left: 4px solid #e67700;
    }
    .dia-titulo-linha {
        font-weight: 700;
        font-size: 0.92rem;
        color: #102a43;
        margin-bottom: 2px;
    }
    .dia-sub-linha {
        font-size: 0.78rem;
        color: #627d98;
    }
    .feriado-selo {
        display: inline-block;
        background-color: #fef3c7;
        color: #9c4221;
        border: 1px solid #fcd34d;
        font-size: 0.74rem;
        font-weight: 700;
        padding: 2px 7px;
        border-radius: 4px;
        margin-top: 4px;
    }

    /* Caixa de Notas à direita */
    .notas-card {
        background-color: #fffdf5;
        border: 1px solid #f0d58c;
        border-left: 5px solid #e67700;
        border-radius: 6px;
        padding: 14px 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-top: 4px;
    }
    .notas-titulo {
        font-size: 0.94rem;
        font-weight: 700;
        color: #9c4221;
        margin-bottom: 10px;
    }
    .feriado-item {
        background: #ffffff;
        border: 1px solid #fae3a8;
        border-radius: 5px;
        padding: 8px 11px;
        margin-bottom: 8px;
        font-size: 0.88rem;
        color: #102a43;
    }
    .feriado-item:last-child {
        margin-bottom: 0;
    }
    .badge-nac {
        display: inline-block;
        background: #e7f5ff;
        color: #1560bd;
        font-size: 0.74rem;
        font-weight: 700;
        padding: 2px 8px;
        border-radius: 10px;
        margin-top: 4px;
    }
    .badge-reg {
        display: inline-block;
        background: #fff4e6;
        color: #d9480f;
        font-size: 0.74rem;
        font-weight: 700;
        padding: 2px 8px;
        border-radius: 10px;
        margin-top: 4px;
    }

    /* Botões */
    div.stButton > button[kind="primary"] {
        background-color: #1560bd !important;
        border-color: #1560bd !important;
        color: #ffffff !important;
        font-weight: 600;
    }
    div.stButton > button[kind="primary"]:hover {
        background-color: #0f4c9a !important;
        border-color: #0f4c9a !important;
    }
    div.stButton > button {
        background-color: #ffffff;
        color: #1f2933;
        border: 1px solid #cbd2d9;
    }

    /* Histórico */
    .hist-card {
        background: #ffffff;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        border: 1px solid #d9e2ec;
        padding: 16px 20px;
        margin-top: 18px;
    }

    @media (max-width: 768px) {
        .block-container {
            padding-top: 3.6rem !important;
            padding-left: 0.8rem !important;
            padding-right: 0.8rem !important;
        }
        .afaplan-banner {
            flex-direction: column;
        }
        .afaplan-logo-box {
            min-width: 100%;
            border-bottom: 1px solid #e4e7eb;
        }
        .afaplan-title-box {
            justify-content: center;
            text-align: center;
            padding: 12px 14px;
        }
        .afaplan-title-box h1 {
            font-size: 1.05rem !important;
        }
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. ESTRUTURA DE PROJETOS, SUBDISCIPLINAS, CREDENCIAIS E COLABORADORES
# ==============================================================================
ESTRUTURA_PADRAO = {
    "2041 - PARÁ": {
        "uf": "PA",
        "login_geral": "lider.pa",
        "senha_geral": "pa@2041",
        "subdisciplinas": {
            "SMC": {
                "login": "smc.pa",
                "senha": "pa@2041",
                "colaboradores": ["Maria Silva", "João Santos", "Ana Costa"]
            },
            "Fiscalização": {
                "login": "fisc.pa",
                "senha": "pa@2041",
                "colaboradores": ["Carlos Souza", "Fernanda Lima", "Paulo Henrique"]
            },
            "Projetos": {
                "login": "proj.pa",
                "senha": "pa@2041",
                "colaboradores": ["Lucas Almeida", "Bruna Rocha", "Diego Ribeiro"]
            },
            "Ambiental": {
                "login": "amb.pa",
                "senha": "pa@2041",
                "colaboradores": ["Camila Ferreira", "Rafael Gomes", "Mariana Dias"]
            }
        }
    },
    "2042 - AMAPÁ": {
        "uf": "AP",
        "login_geral": "lider.ap",
        "senha_geral": "ap@2042",
        "subdisciplinas": {
            "Projetos": {
                "login": "proj.ap",
                "senha": "ap@2042",
                "colaboradores": ["Roberto Alves", "Juliana Mendes", "Thiago Santos"]
            },
            "Ambiental": {
                "login": "amb.ap",
                "senha": "ap@2042",
                "colaboradores": ["Marcos Pereira", "Patrícia Rocha", "Letícia Souza"]
            },
            "Fiscalização": {
                "login": "fisc.ap",
                "senha": "ap@2042",
                "colaboradores": ["Eduardo Lima", "Beatriz Martins", "Gabriel Costa"]
            }
        }
    },
    "2043 - GOIÁS": {
        "uf": "GO",
        "login_geral": "lider.go",
        "senha_geral": "go@2043",
        "subdisciplinas": {
            "Projetos": {
                "login": "proj.go",
                "senha": "go@2043",
                "colaboradores": ["Lucas Oliveira", "Camila Ferreira", "André Barbosa"]
            },
            "Ambiental": {
                "login": "amb.go",
                "senha": "go@2043",
                "colaboradores": ["Rafael Gomes", "Beatriz Martins", "Vanessa Cruz"]
            },
            "Fiscalização": {
                "login": "fisc.go",
                "senha": "go@2043",
                "colaboradores": ["Rodrigo Nunes", "Amanda Silva", "Felipe Castro"]
            }
        }
    }
}

# ==============================================================================
# 3. FERIADOS NACIONAIS E REGIONAIS (PA, AP, GO) + CICLO QUINTA A QUARTA
# ==============================================================================
FERIADOS_CADASTRO = {
    # Nacionais
    (1, 1): {"nome": "Confraternização Universal", "escopo": "Nacional"},
    (4, 21): {"nome": "Tiradentes", "escopo": "Nacional"},
    (5, 1): {"nome": "Dia do Trabalhador", "escopo": "Nacional"},
    (9, 7): {"nome": "Independência do Brasil", "escopo": "Nacional"},
    (10, 12): {"nome": "Nossa Senhora Aparecida", "escopo": "Nacional"},
    (11, 2): {"nome": "Finados", "escopo": "Nacional"},
    (11, 15): {"nome": "Proclamação da República", "escopo": "Nacional"},
    (11, 20): {"nome": "Dia Nacional de Zumbi e da Consciência Negra", "escopo": "Nacional"},
    (12, 25): {"nome": "Natal", "escopo": "Nacional"},

    # Regionais - Pará (PA)
    (1, 7): {"nome": "Dia da Liberdade de Cultos (PA)", "escopo": "PA"},
    (8, 15): {"nome": "Adesão do Pará à Independência do Brasil", "escopo": "PA"},

    # Regionais - Amapá (AP)
    (3, 19): {"nome": "Dia de São José (AP)", "escopo": "AP"},
    (9, 13): {"nome": "Criação do Território Federal do Amapá", "escopo": "AP"},

    # Regionais - Goiás (GO)
    (7, 26): {"nome": "Fundação da Cidade de Goiás (GO)", "escopo": "GO"},
    (10, 24): {"nome": "Pedra Fundamental de Goiânia (GO)", "escopo": "GO"},
}

MESES_PT = {
    1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril",
    5: "Maio", 6: "Junho", 7: "Julho", 8: "Agosto",
    9: "Setembro", 10: "Outubro", 11: "Novembro", 12: "Dezembro"
}

DIAS_SEMANA_PT = {
    0: "Segunda-feira",
    1: "Terça-feira",
    2: "Quarta-feira",
    3: "Quinta-feira",
    4: "Sexta-feira",
    5: "Sábado",
    6: "Domingo"
}

OPCOES_MES_ANO = [
    f"{MESES_PT[m]} / 2026" for m in range(1, 13)
] + [
    f"{MESES_PT[m]} / 2027" for m in range(1, 7)
]


def gerar_semanas_quinta_quarta(mes_ano_str):
    partes = [p.strip() for p in mes_ano_str.split("/")]
    nome_mes = partes[0]
    ano = int(partes[1])
    mes_num = {v: k for k, v in MESES_PT.items()}[nome_mes]

    primeiro_dia = date(ano, mes_num, 1)
    ultimo_dia = date(ano, mes_num, calendar.monthrange(ano, mes_num)[1])

    offset_para_quinta = (primeiro_dia.weekday() - 3) % 7
    inicio_ciclo = primeiro_dia - timedelta(days=offset_para_quinta)

    dias_no_mes = (inicio_ciclo + timedelta(days=6) - primeiro_dia).days + 1
    if inicio_ciclo < primeiro_dia and dias_no_mes < 4:
        inicio_ciclo = inicio_ciclo + timedelta(days=7)

    semanas = []
    idx = 1
    atual_ini = inicio_ciclo
    while atual_ini <= ultimo_dia:
        atual_fim = atual_ini + timedelta(days=6)
        dias_lista = [atual_ini + timedelta(days=i) for i in range(7)]
        semanas.append({
            "chave": f"Semana {idx}",
            "rotulo": f"Semana {idx} ({atual_ini.strftime('%d/%m')} a {atual_fim.strftime('%d/%m')})",
            "inicio": atual_ini,
            "fim": atual_fim,
            "dias": dias_lista
        })
        idx += 1
        atual_ini += timedelta(days=7)

    return semanas


# ==============================================================================
# 4. PERSISTÊNCIA EM JSON E EXCEL (COM BLOQUEIO ANTI-SOBREPOSIÇÃO)
# ==============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_DIR = os.path.join(BASE_DIR, "database")
STATIC_DIR = os.path.join(BASE_DIR, "static")
os.makedirs(DB_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

GRUPOS_JSON_PATH = os.path.join(DB_DIR, "colaboradores_por_grupo.json")
EXCEL_PATH = os.path.join(DB_DIR, "justificativas_afaplan.xlsx")
LOCK_PATH = os.path.join(DB_DIR, "justificativas_afaplan.xlsx.lock")

COLUNAS_EXCEL = [
    "ID_Registro",
    "Data_Hora_Registro",
    "Projeto",
    "Subdisciplina",
    "Usuario_Logado",
    "Colaborador",
    "Mes_Operacao",
    "Semana_Operacao",
    "Justificativa"
]


def localizar_logo_path():
    candidatos = [
        "Logo_afaplan_1.png",
        "Logo_afaplan_1 (1).png",
        "logo_afaplan_1.png",
        "Logo_afaplan_2.png",
        "Logo_afaplan_2 (1).png"
    ]
    for nome in candidatos:
        caminho = os.path.join(STATIC_DIR, nome)
        if os.path.exists(caminho):
            return caminho
    for arq in sorted(os.listdir(STATIC_DIR)):
        if arq.lower().endswith((".png", ".jpg", ".jpeg", ".webp")):
            return os.path.join(STATIC_DIR, arq)
    return None


def carregar_estrutura_grupos():
    if not os.path.exists(GRUPOS_JSON_PATH):
        with open(GRUPOS_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump(ESTRUTURA_PADRAO, f, ensure_ascii=False, indent=4)
        return ESTRUTURA_PADRAO

    with open(GRUPOS_JSON_PATH, "r", encoding="utf-8") as f:
        dados = json.load(f)

    for proj, cfg_padrao in ESTRUTURA_PADRAO.items():
        if proj not in dados:
            dados[proj] = cfg_padrao
        else:
            dados[proj]["uf"] = cfg_padrao["uf"]
            for sub, sub_cfg in cfg_padrao["subdisciplinas"].items():
                if sub not in dados[proj]["subdisciplinas"]:
                    dados[proj]["subdisciplinas"][sub] = sub_cfg
    return dados


def inicializar_excel():
    with FileLock(LOCK_PATH, timeout=15):
        if not os.path.exists(EXCEL_PATH):
            wb = Workbook()
            ws = wb.active
            ws.title = "Base_Justificativas"
            ws.append(COLUNAS_EXCEL)

            header_fill = PatternFill(start_color="1560BD", end_color="1560BD", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)

            for col_num, _ in enumerate(COLUNAS_EXCEL, 1):
                cell = ws.cell(row=1, column=col_num)
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

            larguras = [15, 22, 18, 18, 18, 28, 20, 18, 80]
            for i, largura in enumerate(larguras, 1):
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = largura

            wb.save(EXCEL_PATH)


def ler_registros_excel(projeto_atual, sub_atual):
    inicializar_excel()
    registros = []
    with FileLock(LOCK_PATH, timeout=15):
        wb = load_workbook(EXCEL_PATH, data_only=True)
        ws = wb["Base_Justificativas"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            if str(row[2]) == projeto_atual and str(row[3]) == sub_atual:
                registros.append({
                    "id": str(row[0]),
                    "data_hora": str(row[1]),
                    "projeto": str(row[2]),
                    "subdisciplina": str(row[3]),
                    "usuario": str(row[4]),
                    "colaborador": str(row[5]),
                    "mes": str(row[6]),
                    "semana": str(row[7]),
                    "justificativa": str(row[8])
                })
        wb.close()
    registros.reverse()
    return registros


def salvar_registro_excel(projeto, subdisciplina, usuario, colaborador, mes_op, semana_op, justificativa):
    inicializar_excel()
    with FileLock(LOCK_PATH, timeout=20):
        wb = load_workbook(EXCEL_PATH)
        ws = wb["Base_Justificativas"]
        novo_id = datetime.now().strftime("%Y%m%d%H%M%S%f")[:18]
        data_hora_fmt = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        ws.append([
            novo_id,
            data_hora_fmt,
            projeto,
            subdisciplina,
            usuario,
            colaborador,
            mes_op,
            semana_op,
            justificativa
        ])
        wb.save(EXCEL_PATH)
        wb.close()


def excluir_registro_excel(registro_id, projeto_atual, sub_atual):
    with FileLock(LOCK_PATH, timeout=20):
        wb = load_workbook(EXCEL_PATH)
        ws = wb["Base_Justificativas"]
        linha_del = None
        for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if str(row[0]) == str(registro_id) and str(row[2]) == projeto_atual and str(row[3]) == sub_atual:
                linha_del = idx
                break
        if linha_del:
            ws.delete_rows(linha_del, 1)
            wb.save(EXCEL_PATH)
            wb.close()
            return True
        wb.close()
        return False


inicializar_excel()
grupos_dados = carregar_estrutura_grupos()

# ==============================================================================
# 5. CONTROLE DE SESSÃO (LOGIN / LOGOUT)
# ==============================================================================
if "logado" not in st.session_state:
    st.session_state.logado = False
    st.session_state.projeto = ""
    st.session_state.subdisciplina = ""
    st.session_state.usuario = ""
    st.session_state.uf = "PA"


def render_banner(badge_texto=""):
    badge_html = f'<span class="afaplan-badge">{badge_texto}</span>' if badge_texto else ""
    st.markdown(
        f'<div class="afaplan-banner">'
        f'<div class="afaplan-logo-box">'
        f'<svg width="165" height="42" viewBox="0 0 170 44" xmlns="http://www.w3.org/2000/svg">'
        f'<path d="M8 36 C8 16, 22 8, 40 8 L40 36 Z" fill="none" stroke="#1560bd" stroke-width="6"/>'
        f'<text x="50" y="30" font-family="Segoe UI, Arial, sans-serif" font-size="26" font-weight="700" fill="#2d3136">afaplan</text>'
        f'</svg>'
        f'</div>'
        f'<div class="afaplan-title-box">'
        f'<h1>REGISTRO DE JUSTIFICATIVAS DE DESEMPENHO</h1>'
        f'{badge_html}'
        f'</div>'
        f'</div>',
        unsafe_allow_html=True
    )


def tela_login():
    render_banner()
    _, col_centro, _ = st.columns([1, 1.2, 1])
    with col_centro:
        logo_path = localizar_logo_path()
        if logo_path:
            st.image(logo_path, width=200)
        st.markdown("### Acesso ao Sistema de Justificativas")
        projeto_sel = st.selectbox("1. Selecione o Projeto", list(grupos_dados.keys()))
        subdisciplinas_disp = list(grupos_dados[projeto_sel]["subdisciplinas"].keys())
        sub_sel = st.selectbox("2. Selecione a Subdisciplina", subdisciplinas_disp)

        with st.form("form_login_st"):
            usuario_inp = st.text_input("3. Login de Acesso", placeholder="Digite seu usuário")
            senha_inp = st.text_input("4. Senha", type="password", placeholder="Digite sua senha")
            entrar = st.form_submit_button("Acessar Painel", type="primary", use_container_width=True)

            if entrar:
                proj_cfg = grupos_dados[projeto_sel]
                sub_cfg = proj_cfg["subdisciplinas"][sub_sel]
                u_clean = usuario_inp.strip()

                valido_sub = (u_clean == sub_cfg["login"] and senha_inp == sub_cfg["senha"])
                valido_geral = (u_clean == proj_cfg["login_geral"] and senha_inp == proj_cfg["senha_geral"])

                if valido_sub or valido_geral:
                    st.session_state.logado = True
                    st.session_state.projeto = projeto_sel
                    st.session_state.subdisciplina = sub_sel
                    st.session_state.usuario = u_clean
                    st.session_state.uf = proj_cfg["uf"]
                    st.rerun()
                else:
                    st.error("Projeto, subdisciplina, login ou senha incorretos.")


def tela_registro():
    projeto = st.session_state.projeto
    subdisciplina = st.session_state.subdisciplina
    usuario = st.session_state.usuario
    uf_projeto = st.session_state.uf

    render_banner(f"{projeto} • {subdisciplina} ({usuario})")

    _, col_sair = st.columns([8, 1])
    with col_sair:
        if st.button("Sair", use_container_width=True):
            st.session_state.logado = False
            st.rerun()

    # Linha 1: Colaborador | Mês/Ano | Semana (Quinta a Quarta)
    colaboradores = grupos_dados[projeto]["subdisciplinas"][subdisciplina]["colaboradores"]
    idx_nov = OPCOES_MES_ANO.index("Novembro / 2026") if "Novembro / 2026" in OPCOES_MES_ANO else 0

    c1, c2, c3 = st.columns([1.5, 1.1, 1.4])
    with c1:
        colaborador_sel = st.selectbox("Nome do Colaborador", colaboradores)
    with c2:
        mes_sel = st.selectbox("Mês/Ano de Operação", OPCOES_MES_ANO, index=idx_nov)
    with c3:
        semanas_mes = gerar_semanas_quinta_quarta(mes_sel)
        rotulos_semanas = [s["rotulo"] for s in semanas_mes]
        semana_rotulo_sel = st.selectbox("Semana de Operação (Quinta a Quarta)", rotulos_semanas)

    semana_obj = next(s for s in semanas_mes if s["rotulo"] == semana_rotulo_sel)
    semana_chave = semana_obj["chave"]

    # Linha 2: Justificativa Diária Centralizada/Alinhada (Esquerda) + Notas de Feriados (Direita)
    col_dias, col_notas = st.columns([1.85, 1])

    feriados_semana = []
    with col_dias:
        st.markdown("**Descrição Diária dos Acontecimentos da Semana (Justificativa por dia)**")
        st.caption(
            f"Período da Semana: {semana_obj['inicio'].strftime('%d/%m/%Y')} (Quinta) a "
            f"{semana_obj['fim'].strftime('%d/%m/%Y')} (Quarta)  •  Ciclo: Quinta-feira a Quarta-feira"
        )

        entradas_dias = []
        for d in semana_obj["dias"]:
            info_feriado = FERIADOS_CADASTRO.get((d.month, d.day))
            feriado_valido = None
            if info_feriado and info_feriado["escopo"] in ("Nacional", uf_projeto):
                tipo_lbl = (
                    "Feriado Nacional"
                    if info_feriado["escopo"] == "Nacional"
                    else f"Feriado Regional ({info_feriado['escopo']})"
                )
                feriado_valido = {
                    "data": d.strftime("%d/%m/%Y"),
                    "dia_extenso": f"{d.day:02d} de {MESES_PT[d.month].lower()}",
                    "nome": info_feriado["nome"],
                    "tipo": tipo_lbl
                }
                feriados_semana.append(feriado_valido)

            data_curta = d.strftime("%d/%m")
            data_comp = d.strftime("%d/%m/%Y")
            dia_sem_nome = DIAS_SEMANA_PT[d.weekday()]
            key_input = f"inp_{mes_sel}_{semana_chave}_{data_curta}"

            box_cls = "dia-card-box feriado-destaque" if feriado_valido else "dia-card-box"
            selo_html = (
                f'<div><span class="feriado-selo">Feriado: {feriado_valido["nome"]}</span></div>'
                if feriado_valido else ""
            )

            col_lbl, col_inp = st.columns([1.05, 2.2], vertical_alignment="center")
            with col_lbl:
                st.markdown(
                    f'<div class="{box_cls}">'
                    f'<div class="dia-titulo-linha">{data_curta} — {dia_sem_nome}</div>'
                    f'<div class="dia-sub-linha">{data_comp}</div>'
                    f'{selo_html}'
                    f'</div>',
                    unsafe_allow_html=True
                )
            with col_inp:
                placeholder_txt = (
                    f"Ex.: Feriado ({feriado_valido['nome']}) ou ocorrência do dia {data_curta}..."
                    if feriado_valido
                    else f"Ocorrência do dia {data_curta} (Ex.: Fiquei doente, Atestado médico, Não se aplica...)"
                )
                texto_dia = st.text_input(
                    f"Justificativa {data_curta}",
                    key=key_input,
                    placeholder=placeholder_txt,
                    label_visibility="collapsed"
                )
            entradas_dias.append((data_curta, texto_dia))

    with col_notas:
        st.markdown("**Notas**")
        if not feriados_semana:
            notas_html = (
                f'<div class="notas-card">'
                f'<div class="notas-titulo">Feriados da Semana ({uf_projeto})</div>'
                f'<div style="color:#627d98; font-style:italic; font-size:0.88rem;">'
                f'Nenhum feriado nacional ou regional ({uf_projeto}) identificado para {semana_chave.lower()} de {mes_sel.lower()}.'
                f'</div></div>'
            )
        else:
            itens_parts = []
            for f in feriados_semana:
                badge_cls = "badge-nac" if "Nacional" in f["tipo"] else "badge-reg"
                itens_parts.append(
                    f'<div class="feriado-item">'
                    f'<div><strong>{f["dia_extenso"]} ({f["data"]}) - {f["nome"]}</strong></div>'
                    f'<span class="{badge_cls}">{f["tipo"]}</span>'
                    f'</div>'
                )
            notas_html = (
                f'<div class="notas-card">'
                f'<div class="notas-titulo">Feriados da Semana ({uf_projeto})</div>'
                + "".join(itens_parts)
                + '</div>'
            )
        st.markdown(notas_html, unsafe_allow_html=True)

    # Botões de Ação centralizados (Cancelar | Limpar | Cadastrar)
    st.markdown("<br>", unsafe_allow_html=True)
    _, b1, b2, b3, _ = st.columns([1.8, 1, 1, 1, 1.8])

    with b1:
        if st.button("Cancelar", use_container_width=True):
            st.session_state.logado = False
            st.rerun()

    with b2:
        if st.button("Limpar", use_container_width=True):
            for d in semana_obj["dias"]:
                k = f"inp_{mes_sel}_{semana_chave}_{d.strftime('%d/%m')}"
                if k in st.session_state:
                    st.session_state[k] = ""
            st.rerun()

    with b3:
        if st.button("Cadastrar", type="primary", use_container_width=True):
            partes = [f"{dt_c} - {txt.strip()}" for dt_c, txt in entradas_dias if txt.strip()]
            if not partes:
                st.warning("Preencha a justificativa de pelo menos um dia da semana antes de cadastrar.")
            else:
                justificativa_somada = "; ".join(partes) + "."
                salvar_registro_excel(
                    projeto=projeto,
                    subdisciplina=subdisciplina,
                    usuario=usuario,
                    colaborador=colaborador_sel,
                    mes_op=mes_sel,
                    semana_op=semana_chave,
                    justificativa=justificativa_somada
                )
                for d in semana_obj["dias"]:
                    k = f"inp_{mes_sel}_{semana_chave}_{d.strftime('%d/%m')}"
                    if k in st.session_state:
                        st.session_state[k] = ""
                st.success("Justificativa semanal consolidada e registrada no Excel com sucesso!")
                st.rerun()

    # Histórico de Registros + Download do Excel
    st.markdown('<div class="hist-card">', unsafe_allow_html=True)
    col_hist_tit, col_dl = st.columns([4, 1.3])
    with col_hist_tit:
        st.markdown("### Histórico de Justificativas")
    with col_dl:
        if os.path.exists(EXCEL_PATH):
            with open(EXCEL_PATH, "rb") as f_xlsx:
                st.download_button(
                    label="Baixar Excel (.xlsx)",
                    data=f_xlsx.read(),
                    file_name="justificativas_afaplan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    registros = ler_registros_excel(projeto, subdisciplina)
    if not registros:
        st.info("Nenhuma justificativa registrada para este projeto e subdisciplina ainda.")
    else:
        h1, h2, h3, h4, h5, h6 = st.columns([1.2, 1.3, 1.1, 0.8, 3.2, 0.7])
        h1.markdown("**Data/Hora**")
        h2.markdown("**Colaborador**")
        h3.markdown("**Mês de Operação**")
        h4.markdown("**Semana**")
        h5.markdown("**Descrição da Semana**")
        h6.markdown("**Ações**")
        st.markdown("<hr style='margin: 4px 0 10px 0;'>", unsafe_allow_html=True)

        for reg in registros:
            r1, r2, r3, r4, r5, r6 = st.columns([1.2, 1.3, 1.1, 0.8, 3.2, 0.7])
            r1.write(reg["data_hora"])
            r2.markdown(f"**{reg['colaborador']}**")
            r3.write(reg["mes"])
            r4.write(reg["semana"])
            r5.write(reg["justificativa"])
            if r6.button("Excluir", key=f"del_{reg['id']}"):
                excluir_registro_excel(reg["id"], projeto, subdisciplina)
                st.rerun()
            st.markdown("<hr style='margin: 6px 0; border-top: 1px solid #e4e7eb;'>", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)


if st.session_state.logado:
    tela_registro()
else:
    tela_login()
