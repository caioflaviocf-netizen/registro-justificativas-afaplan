import os
import json
import calendar
from datetime import datetime, date, timedelta, timezone
try:
    from zoneinfo import ZoneInfo
    TZ_BRASILIA = ZoneInfo("America/Sao_Paulo")
except Exception:
    TZ_BRASILIA = timezone(timedelta(hours=-3))

def agora_brasilia():
    return datetime.now(TZ_BRASILIA)
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
        font-size: 0.94rem !important;
    }

    /* Melhora a nitidez e o contraste do texto modelo (placeholder) dentro das caixas de texto */
    input::placeholder,
    textarea::placeholder,
    div[data-baseweb="input"] input::placeholder {
        color: #486581 !important;
        opacity: 1 !important;
        font-size: 0.88rem !important;
        font-weight: 400 !important;
    }

    /* Card de cada linha diária */
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
        margin-bottom: 16px;
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

    /* Card de Histórico de Acessos na coluna direita */
    .acessos-card {
        background-color: #ffffff;
        border: 1px solid #d9e2ec;
        border-left: 5px solid #1560bd;
        border-radius: 6px;
        padding: 14px 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-top: 4px;
    }
    .acessos-titulo {
        font-size: 0.94rem;
        font-weight: 700;
        color: #102a43;
        margin-bottom: 10px;
    }
    .acesso-item {
        background: #f8fafc;
        border: 1px solid #e4e7eb;
        border-radius: 5px;
        padding: 8px 10px;
        margin-bottom: 8px;
        font-size: 0.84rem;
        color: #334e68;
    }
    .acesso-item:last-child {
        margin-bottom: 0;
    }
    .acesso-hora {
        font-size: 0.76rem;
        font-weight: 700;
        color: #1560bd;
    }
    .acesso-vazio {
        color: #627d98;
        font-style: italic;
        font-size: 0.85rem;
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

    /* Histórico de Justificativas */
    .hist-card {
        background: #ffffff;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        border: 1px solid #d9e2ec;
        padding: 16px 20px;
        margin-top: 18px;
    }
    .hist-vazio-box {
        background-color: #f8fafc;
        border: 1px solid #e4e7eb;
        border-radius: 6px;
        padding: 12px 16px;
        color: #486581;
        font-size: 0.90rem;
        text-align: center;
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
# 2. CADASTRO DE PROJETOS, LÍDERES E COLABORADORES ATIVOS (BASE PASTA1.XLSX)
# ==============================================================================
ACESSOS_PROJETOS = {
    "2041 - PARÁ": {
        "uf": "PA",
        "lider": {
            "login": "lider.pa",
            "senha": "pa@2041"
        },
        "usuarios": {
            "adivaldo.santos": {
                "nome": "Adivaldo Gabriel Albuquerque dos Santos",
                "senha": "AS@0718",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "adryano.bastos": {
                "nome": "Adryano Miranda Bastos",
                "senha": "AB@1398",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "alan.santos": {
                "nome": "Alan Robin Bezerra dos Santos",
                "senha": "AS@1445",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "alana.menezes": {
                "nome": "Alana Cristina Rocha Menezes",
                "senha": "AM@8902",
                "funcao": "Técnico(a) de Meio Ambiente Junior, I"
            },
            "alessandra.feitosa": {
                "nome": "Alessandra Souza Feitosa",
                "senha": "AF@0507",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "aline.gomes": {
                "nome": "Aline Nogueira Gomes",
                "senha": "AG@7540",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "anderson.matos": {
                "nome": "Anderson Lucas Lopes de Matos",
                "senha": "AM@4577",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "anderson.mendes": {
                "nome": "Anderson dos Santos Mendes",
                "senha": "AM@2153",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "antonio.silva": {
                "nome": "Antonio Aldair Oliveira da Silva",
                "senha": "AS@7249",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "antonio.calheiros": {
                "nome": "Antonio Augusto Andrade Calheiros",
                "senha": "AC@4611",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "antonio.correa": {
                "nome": "Antonio Luis Carneiro Correa",
                "senha": "AC@0451",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "arilene.pacheco": {
                "nome": "Arilene de Fatima de Amorim Pacheco",
                "senha": "AP@8516",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "benedito.leite": {
                "nome": "Benedito Assuncao Moraes Leite",
                "senha": "BL@3487",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "carlos.junior": {
                "nome": "Carlos Alberto Brasil Ferreira Junior",
                "senha": "CJ@2580",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "carlos.pessoa": {
                "nome": "Carlos Andre Goncalves Pessoa",
                "senha": "CP@6460",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "carlos.silva": {
                "nome": "Carlos Augusto Silva da Silva",
                "senha": "CS@8446",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "carlos.luz": {
                "nome": "Carlos Emanuel Martins da Luz",
                "senha": "CL@4608",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "carlos.alcantara": {
                "nome": "Carlos Vinicius Silva de Alcantara",
                "senha": "CA@4408",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "carlos.junior2": {
                "nome": "Carlos dos Reis Costa Junior",
                "senha": "CJ@1499",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "claudia.guerra": {
                "nome": "Claudia Gabriela Souza Guerra",
                "senha": "CG@2780",
                "funcao": "Jovem Aprendiz, Geral"
            },
            "cledison.oliveira": {
                "nome": "Cledison Ferreira de Oliveira",
                "senha": "CO@3650",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "dayves.vieira": {
                "nome": "Dayves Raphael Torres Vieira",
                "senha": "DV@2894",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "diego.ferreira": {
                "nome": "Diego Moraes Ferreira",
                "senha": "DF@8837",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "diogo.ferreira": {
                "nome": "Diogo Moraes Ferreira",
                "senha": "DF@6518",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "edilson.santos": {
                "nome": "Edilson Pantoja dos Santos",
                "senha": "ES@8556",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "elizeu.silva": {
                "nome": "Elizeu Rocha da Silva",
                "senha": "ES@4619",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "emerson.silva": {
                "nome": "Emerson Roberto Martins da Silva",
                "senha": "ES@4415",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "erick.guedes": {
                "nome": "Erick Cristiano Silva Guedes",
                "senha": "EG@5396",
                "funcao": "Coordenador(a) de Obras Junior, I"
            },
            "eridalton.santos": {
                "nome": "Eridalton da Costa Santos",
                "senha": "ES@5207",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "erika.lopes": {
                "nome": "Erika Sophia Silva Lopes",
                "senha": "EL@0550",
                "funcao": "Assistente Administrativo Pleno, V"
            },
            "erisson.lopes": {
                "nome": "Erisson George Evaristo Lopes",
                "senha": "EL@4648",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "evenilton.moraes": {
                "nome": "Evenilton Max Evangelista Moraes",
                "senha": "EM@3729",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "everton.marinho": {
                "nome": "Everton Thiago Alcantara Marinho",
                "senha": "EM@6728",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "fabio.barros": {
                "nome": "Fabio Bruno Medeiros de Barros",
                "senha": "FB@0894",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "fabio.pereira": {
                "nome": "Fabio Correa Pereira",
                "senha": "FP@1106",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "fagner.silva": {
                "nome": "Fagner Freitas da Silva",
                "senha": "FS@7556",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "fernanda.pina": {
                "nome": "Fernanda Gabryela Costa Pina",
                "senha": "FP@4615",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "fernanda.loureiro": {
                "nome": "Fernanda de Melo Loureiro",
                "senha": "FL@0910",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "franklin.santos": {
                "nome": "Franklin Costa dos Santos",
                "senha": "FS@8102",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "fredson.santos": {
                "nome": "Fredson Santos dos Santos",
                "senha": "FS@0912",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "gabriel.santos": {
                "nome": "Gabriel Pereira dos Santos",
                "senha": "GS@0427",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "gedielson.silva": {
                "nome": "Gedielson Oliveira da Silva",
                "senha": "GS@5056",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "geovana.ferreira": {
                "nome": "Geovana Neves Ferreira",
                "senha": "GF@4497",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "gilcelio.silveira": {
                "nome": "Gilcelio Lopes Silveira",
                "senha": "GS@2974",
                "funcao": "PS - Coordenador(a) Eletricista, PJ"
            },
            "glauberto.lima": {
                "nome": "Glauberto Pinheiro Lima",
                "senha": "GL@8779",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "gleiciane.correa": {
                "nome": "Gleiciane da Silva Correa",
                "senha": "GC@4634",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "helio.souza": {
                "nome": "Helio Mario Cezar Souza",
                "senha": "HS@4376",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "heliton.nunes": {
                "nome": "Heliton Oliveira Nunes",
                "senha": "HN@5175",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "henderson.santos": {
                "nome": "Henderson Duarte dos Santos",
                "senha": "HS@6813",
                "funcao": "PS - Engenheiro(a) Florestal, PJ"
            },
            "henrique.oliveira": {
                "nome": "Henrique Cezar Pimentel Oliveira",
                "senha": "HO@4269",
                "funcao": "Técnico(a) Eletrotécnica Pleno, V"
            },
            "iago.cerbino": {
                "nome": "Iago Fernando Duarte Cerbino",
                "senha": "IC@1941",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "jaderson.oliveira": {
                "nome": "Jaderson dos Remedios Oliveira",
                "senha": "JO@4252",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "jamilson.correa": {
                "nome": "Jamilson Ferreira Correa",
                "senha": "JC@3101",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "jhon.lima": {
                "nome": "Jhon Kennedy Andrade Lima",
                "senha": "JL@2241",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "jhonatan.silva": {
                "nome": "Jhonatan Deyvison Dantas Silva",
                "senha": "JS@4152",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "jhonatha.santos": {
                "nome": "Jhonatha Portal dos Santos",
                "senha": "JS@4602",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "joao.brasil": {
                "nome": "Joao Calebe Silva Lima Brasil",
                "senha": "JB@3449",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "joao.almeida": {
                "nome": "Joao Paulo dos santos de Almeida",
                "senha": "JA@5170",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "joaquim.brito": {
                "nome": "Joaquim de Almeida Brito",
                "senha": "JB@7484",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "john.cardoso": {
                "nome": "John Breno da Costa Cardoso",
                "senha": "JC@0516",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "johny.nascimento": {
                "nome": "Johny Tiago Martins Nascimento",
                "senha": "JN@3213",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "jose.carvalho": {
                "nome": "Jose Alberto Monteiro Carvalho",
                "senha": "JC@2893",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "jose.cabral": {
                "nome": "Jose Otavio Rodrigues Cabral",
                "senha": "JC@4261",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "josiane.araujo": {
                "nome": "Josiane Progenio Araujo",
                "senha": "JA@0682",
                "funcao": "Supervisor(a) Administrativo Junior, Geral"
            },
            "josiel.carvalho": {
                "nome": "Josiel dos Santos Carvalho",
                "senha": "JC@9076",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "josimar.souza": {
                "nome": "Josimar Costa de Souza",
                "senha": "JS@4217",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "jozivaldo.pacheco": {
                "nome": "Jozivaldo do Carmo Pacheco",
                "senha": "JP@0791",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "juliana.pinto": {
                "nome": "Juliana Teixeira Pinto",
                "senha": "JP@3813",
                "funcao": "Técnico(a) de Meio Ambiente Junior, I"
            },
            "keci.assis": {
                "nome": "Keci Jhones Silva de Assis",
                "senha": "KA@7529",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "kellen.moraes": {
                "nome": "Kellen Melissa Ferreira Moraes",
                "senha": "KM@2235",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "lauro.silva": {
                "nome": "Lauro Pinto da Silva",
                "senha": "LS@6443",
                "funcao": "Técnico(a) Eletrotécnica Pleno, VIII"
            },
            "luana.costa": {
                "nome": "Luana Giovana Cunha Costa",
                "senha": "LC@8719",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "lucas.bento": {
                "nome": "Lucas Lourran Oliveira Bento",
                "senha": "LB@2214",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "lucas.aguiar": {
                "nome": "Lucas Silva Aguiar",
                "senha": "LA@3995",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "luis.machado": {
                "nome": "Luis Fernando Correa Machado",
                "senha": "LM@1713",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "luiz.santos": {
                "nome": "Luiz Eduardo Santos dos Santos",
                "senha": "LS@4996",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "luziel.costa": {
                "nome": "Luziel Ataide da Costa",
                "senha": "LC@0955",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "manasses.junior": {
                "nome": "Manasses Elias Pereira Junior",
                "senha": "MJ@3077",
                "funcao": "PJ Gerente de Projetos, Geral"
            },
            "marcelo.brabo": {
                "nome": "Marcelo Albano Brabo",
                "senha": "MB@7548",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "marcelo.santos": {
                "nome": "Marcelo Correa dos Santos",
                "senha": "MS@6459",
                "funcao": "PS - Coordenador(a) Ambiental, PJ"
            },
            "marcelo.junior": {
                "nome": "Marcelo dos Santos Fonseca Junior",
                "senha": "MJ@6665",
                "funcao": "Técnico(a) de Meio Ambiente Junior, I"
            },
            "marcio.sampaio": {
                "nome": "Marcio Brener Pereira Sampaio",
                "senha": "MS@7401",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "marcos.baima": {
                "nome": "Marcos de Souza Baima",
                "senha": "MB@0072",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "marlon.silva": {
                "nome": "Marlon Torres e Silva",
                "senha": "MS@0994",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "mateus.silva": {
                "nome": "Mateus Augusto da Costa Silva",
                "senha": "MS@8869",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "mauro.junior": {
                "nome": "Mauro Pereira Forte Junior",
                "senha": "MJ@4884",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "mauro.peniche": {
                "nome": "Mauro de Oliveira Peniche",
                "senha": "MP@7290",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "mellissa.santos": {
                "nome": "Mellissa Sandhiely Ramos dos Santos",
                "senha": "MS@6515",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "miguel.reis": {
                "nome": "Miguel Miranda dos Reis",
                "senha": "MR@8041",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "nalbert.claudio": {
                "nome": "Nalbert Menezes Claudio",
                "senha": "NC@1771",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "nilson.nascimento": {
                "nome": "Nilson Ubirajara Dias Nascimento",
                "senha": "NN@0139",
                "funcao": "Técnico(a) de Meio Ambiente Pleno, V"
            },
            "osvaldo.neto": {
                "nome": "Osvaldo Albuquerque Menezes Neto",
                "senha": "ON@8579",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "oziel.monteiro": {
                "nome": "Oziel Soares Monteiro",
                "senha": "OM@3783",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "paulo.delgado": {
                "nome": "Paulo Sergio Monteiro Delgado",
                "senha": "PD@8823",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "paulo.silva": {
                "nome": "Paulo Silva da Silva",
                "senha": "PS@1125",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "raimundo.oliveira": {
                "nome": "Raimundo Amaral de Oliveira",
                "senha": "RO@5169",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "raul.teixeira": {
                "nome": "Raul Sampaio Teixeira",
                "senha": "RT@7147",
                "funcao": "Coordenador(a) de Projetos Senior, XII"
            },
            "renataly.silva": {
                "nome": "Renataly Trindade Silva",
                "senha": "RS@6527",
                "funcao": "Técnico(a) de Meio Ambiente Junior, I"
            },
            "ricardo.rego": {
                "nome": "Ricardo Souza do Rego",
                "senha": "RR@0746",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "richard.cardoso": {
                "nome": "Richard Miguel Chagas Cardoso",
                "senha": "RC@6523",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "ricsant.santos": {
                "nome": "Ricsant da Silva Santos",
                "senha": "RS@3075",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "samuel.santos": {
                "nome": "Samuel da Silva dos Santos",
                "senha": "SS@4061",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "samuel.silva": {
                "nome": "Samuel dos Santos Silva",
                "senha": "SS@7570",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "sheyla.oliveira": {
                "nome": "Sheyla Fialho de Oliveira",
                "senha": "SO@0836",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "silvana.oliveira": {
                "nome": "Silvana Lustoza de Oliveira",
                "senha": "SO@8620",
                "funcao": "Analista Administrativo Senior, VIII"
            },
            "silvio.souza": {
                "nome": "Silvio Adriano de Souza",
                "senha": "SS@0845",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "thalis.lima": {
                "nome": "Thalis Efraim Barbosa Lima",
                "senha": "TL@8588",
                "funcao": "Analista Administrativo Senior, VIII"
            },
            "vanda.tembe": {
                "nome": "Vanda Lucia Martins Farias Tembe",
                "senha": "VT@3697",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "victor.wanzeller": {
                "nome": "Victor de Figueiredo Wanzeller",
                "senha": "VW@8616",
                "funcao": "Analista Administrativo Pleno, VII"
            },
            "vitoria.milhomem": {
                "nome": "Vitoria Lancaster Alves Milhomem",
                "senha": "VM@0679",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "wallison.marques": {
                "nome": "Wallison Williams Coimbra Marques",
                "senha": "WM@1317",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "wellington.barros": {
                "nome": "Wellington Carlos Moraes Barros",
                "senha": "WB@3753",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "yago.feitosa": {
                "nome": "Yago Raphael Maia Feitosa",
                "senha": "YF@2148",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "yasmim.sousa": {
                "nome": "Yasmim Costa Sousa",
                "senha": "YS@8975",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "yuri.teles": {
                "nome": "Yuri de Sousa Teles",
                "senha": "YT@2945",
                "funcao": "Técnico(a) em Segurança do Trabalho Junior, I"
            },
            "zidane.lopes": {
                "nome": "Zidane da Silva Lopes",
                "senha": "ZL@1265",
                "funcao": "Apontador(a) de Produção Júnior, I"
            }
        }
    },
    "2042 - AMAPÁ": {
        "uf": "AP",
        "lider": {
            "login": "lider.ap",
            "senha": "ap@2042"
        },
        "usuarios": {
            "alrivaldo.junior": {
                "nome": "Alrivaldo Marques Leal Junior",
                "senha": "AJ@4447",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "andre.jacaranda": {
                "nome": "Andre Luiz da Silva Jacaranda",
                "senha": "AJ@0478",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "bruno.rodrigues": {
                "nome": "Bruno Barbosa Rodrigues",
                "senha": "BR@3740",
                "funcao": "Coordenador(a) de Projetos Junior, I"
            },
            "claudmilson.sales": {
                "nome": "Claudmilson da Silva Sales",
                "senha": "CS@1122",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "gabriel.silva": {
                "nome": "Gabriel Lima da Silva",
                "senha": "GS@6950",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "gustavo.nascimento": {
                "nome": "Gustavo Castilo Nascimento",
                "senha": "GN@4416",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "iara.ferreira": {
                "nome": "Iara Melissa Barbosa Ferreira",
                "senha": "IF@6445",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "ivaldo.magno": {
                "nome": "Ivaldo Lima Magno",
                "senha": "IM@5185",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "jhonny.costa": {
                "nome": "Jhonny William Sa da Costa",
                "senha": "JC@0522",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "jivagno.matos": {
                "nome": "Jivagno Brito Matos",
                "senha": "JM@6397",
                "funcao": "Técnico(a) de Meio Ambiente Junior, I"
            },
            "josiane.prazeres": {
                "nome": "Josiane Monteiro dos Prazeres",
                "senha": "JP@8522",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "leandro.silva": {
                "nome": "Leandro Rodrigues da Silva",
                "senha": "LS@5999",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "marcelo.almeida": {
                "nome": "Marcelo Tavares Almeida",
                "senha": "MA@2976",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "nelson.lemos": {
                "nome": "Nelson Pereira Lemos",
                "senha": "NL@8779",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "paulo.guedes": {
                "nome": "Paulo Sandro Ramos Guedes",
                "senha": "PG@7297",
                "funcao": "Apontador(a) de Produção Júnior, I"
            },
            "paulo.barbosa": {
                "nome": "Paulo Sergio de Oliveira Barbosa",
                "senha": "PB@0359",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "samuel.silva2": {
                "nome": "Samuel Bergh Pereira da Silva",
                "senha": "SS@5827",
                "funcao": "Jovem Aprendiz, Geral"
            },
            "tabita.farias": {
                "nome": "Tabita Moura de Farias",
                "senha": "TF@1091",
                "funcao": "Técnico(a) de Meio Ambiente Pleno, VIII"
            },
            "wilcles.alcantara": {
                "nome": "Wilcles Borges de Alcantara",
                "senha": "WA@7822",
                "funcao": "Analista Administrativo Junior, IV"
            }
        }
    },
    "2043 - GOIÁS": {
        "uf": "GO",
        "lider": {
            "login": "lider.go",
            "senha": "go@2043"
        },
        "usuarios": {
            "admilson.moura": {
                "nome": "Admilson Tertuliano Moura",
                "senha": "AM@9072",
                "funcao": "Técnico(a) Eletrotécnica Pleno, V"
            },
            "alessandro.pinto": {
                "nome": "Alessandro Mendes Pinto",
                "senha": "AP@2881",
                "funcao": "Técnico(a) Eletrotécnica Pleno, V"
            },
            "alex.silva": {
                "nome": "Alex Cleison Ferreira da Silva",
                "senha": "AS@5498",
                "funcao": "Apontador(a) de Produção Pleno, VII"
            },
            "alicio.neto": {
                "nome": "Alicio Costa Neto",
                "senha": "AN@0850",
                "funcao": "Supervisor(a) Eletricista Junior, I"
            },
            "amanda.machado": {
                "nome": "Amanda Silva Reis Machado",
                "senha": "AM@2155",
                "funcao": "Analista Administrativo Pleno, VI"
            },
            "ana.rocha": {
                "nome": "Ana Carolina Lima Rocha",
                "senha": "AR@5323",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "anderson.vieira": {
                "nome": "Anderson Davi Silva Vieira",
                "senha": "AV@1341",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "andressa.santos": {
                "nome": "Andressa Gomes dos Santos",
                "senha": "AS@0941",
                "funcao": "Analista Administrativo Junior, I"
            },
            "anna.lopes": {
                "nome": "Anna Clara Miranda Lopes",
                "senha": "AL@9766",
                "funcao": "Jovem Aprendiz, Geral"
            },
            "antonio.oliveira": {
                "nome": "Antonio Matheus Souza de Oliveira",
                "senha": "AO@8643",
                "funcao": "PS - Engenheiro(a) de Projetos, Geral"
            },
            "ariel.oliveira": {
                "nome": "Ariel Alcantara de Oliveira",
                "senha": "AO@5878",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "atila.anjos": {
                "nome": "Atila Junio Santos Anjos",
                "senha": "AA@0786",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "beatriz.paula": {
                "nome": "Beatriz Castro de Paula",
                "senha": "BP@4969",
                "funcao": "Jovem Aprendiz, Geral"
            },
            "bruno.santos": {
                "nome": "Bruno Jose de Ataide Santos",
                "senha": "BS@1149",
                "funcao": "Técnico(a) Eletrotécnica Pleno, V"
            },
            "bruno.nascimento": {
                "nome": "Bruno Kyoshi Silva Nascimento",
                "senha": "BN@0920",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "carlos.santos": {
                "nome": "Carlos Alves dos Santos",
                "senha": "CS@3585",
                "funcao": "Técnico(a) Eletrotécnica Pleno, V"
            },
            "clarissa.cipriano": {
                "nome": "Clarissa Porto Cipriano",
                "senha": "CC@9169",
                "funcao": "Analista Administrativo Senior, XII"
            },
            "deivid.nogueira": {
                "nome": "Deivid Santos Nogueira",
                "senha": "DN@5393",
                "funcao": "Apontador(a) de Produção Pleno, VII"
            },
            "deivid.resende": {
                "nome": "Deivid da Silva Resende",
                "senha": "DR@1981",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "douglas.rocha": {
                "nome": "Douglas Leonardo Toledo Rocha",
                "senha": "DR@4405",
                "funcao": "Supervisor(a) Eletricista Junior, I"
            },
            "ediemerson.santiago": {
                "nome": "Ediemerson Santos Santiago",
                "senha": "ES@7672",
                "funcao": "Apontador(a) de Produção Pleno, V"
            },
            "eduardo.moura": {
                "nome": "Eduardo Alves de Moura",
                "senha": "EM@0523",
                "funcao": "Técnico(a) Eletrotécnica Pleno, VII"
            },
            "elto.alves": {
                "nome": "Elto Gediel Gramacho Alves",
                "senha": "EA@1566",
                "funcao": "Técnico(a) Eletrotécnica Pleno, V"
            },
            "fabiano.nicoli": {
                "nome": "Fabiano Oliveira Silva Nicoli",
                "senha": "FN@5219",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "gabriel.silva2": {
                "nome": "Gabriel Rezende Silva",
                "senha": "GS@0885",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "geovanna.assuncao": {
                "nome": "Geovanna Meireles Silva Assunção",
                "senha": "GA@1034",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "giovane.silva": {
                "nome": "Giovane Jose da Silva",
                "senha": "GS@8904",
                "funcao": "Apontador(a) de Produção Pleno, V"
            },
            "gleison.marques": {
                "nome": "Gleison Santiago Marques",
                "senha": "GM@2098",
                "funcao": "Técnico(a) Eletrotécnica Pleno, VIII"
            },
            "grazielly.santos": {
                "nome": "Grazielly Rodrigues dos Santos",
                "senha": "GS@0842",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "guilherme.rosa": {
                "nome": "Guilherme de Lima Rosa",
                "senha": "GR@1060",
                "funcao": "Técnico(a) Eletrotécnica Junior, I"
            },
            "hugo.veras": {
                "nome": "Hugo Soares Veras",
                "senha": "HV@9940",
                "funcao": "Supervisor(a) de Projetos Junior, II"
            },
            "iasmim.silva": {
                "nome": "Iasmim Borges da Silva",
                "senha": "IS@0453",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "jean.silva": {
                "nome": "Jean Carlos Soares da Silva",
                "senha": "JS@3817",
                "funcao": "Apontador(a) de Produção Pleno, V"
            },
            "jeremias.santos": {
                "nome": "Jeremias Vitor Suaris dos Santos",
                "senha": "JS@3778",
                "funcao": "Apontador(a) de Produção Pleno, V"
            },
            "joao.pereira": {
                "nome": "Joao Pedro de Freitas Pereira",
                "senha": "JP@3841",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "joao.rosa": {
                "nome": "Joao Victor Teixeira do Prado Rosa",
                "senha": "JR@5961",
                "funcao": "Apontador(a) de Produção Pleno, V"
            },
            "jocasta.prado": {
                "nome": "Jocasta Souza Prado",
                "senha": "JP@2048",
                "funcao": "PS - Engenheiro(a) Florestal, PJ"
            },
            "john.santos": {
                "nome": "John Lennon Alencar dos Santos",
                "senha": "JS@4570",
                "funcao": "Técnico(a) Eletrotécnica Pleno, V"
            },
            "jonata.assuncao": {
                "nome": "Jonata Alves Assuncao",
                "senha": "JA@8000",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "josailton.santos": {
                "nome": "Josailton Guimaraes dos Santos",
                "senha": "JS@1020",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "jose.oliveira": {
                "nome": "Jose Vieira Oliveira",
                "senha": "JO@4380",
                "funcao": "Técnico(a) Eletrotécnica Pleno, V"
            },
            "josiane.dias": {
                "nome": "Josiane Cristina Amaral Dias",
                "senha": "JD@1910",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "juliana.almeida": {
                "nome": "Juliana Farias de Almeida",
                "senha": "JA@0972",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "juraci.neto": {
                "nome": "Juraci Almeida Polvora Neto",
                "senha": "JN@2657",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "karine.costa": {
                "nome": "Karine Silva da Costa",
                "senha": "KC@9532",
                "funcao": "Geógrafo(a), Júnior"
            },
            "larissa.souza": {
                "nome": "Larissa Monielle Moreira de Souza",
                "senha": "LS@0801",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "lorraina.santos": {
                "nome": "Lorraina Vitoria Costa dos Santos",
                "senha": "LS@4300",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "luan.silveira": {
                "nome": "Luan Sousa Silveira",
                "senha": "LS@3625",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "lucas.brito": {
                "nome": "Lucas Barbosa Brito",
                "senha": "LB@7800",
                "funcao": "PS - Engenheiro(a) de Projetos, Geral"
            },
            "lucas.silva": {
                "nome": "Lucas Leonardo Ribeiro e Silva",
                "senha": "LS@1254",
                "funcao": "Técnico(a) de Meio Ambiente Junior, I"
            },
            "lucas.souza": {
                "nome": "Lucas Stheferson Feitosa de Souza",
                "senha": "LS@7025",
                "funcao": "Apontador(a) de Produção Júnior, IV"
            },
            "luiz.cardoso": {
                "nome": "Luiz Daniel Rocha Cardoso",
                "senha": "LC@3849",
                "funcao": "Apontador(a) de Produção Júnior, IV"
            },
            "marcelo.almeida2": {
                "nome": "Marcelo Rodrigues de Almeida",
                "senha": "MA@2877",
                "funcao": "Técnico(a) Eletrotécnica Pleno, V"
            },
            "marcos.fogaca": {
                "nome": "Marcos Antonio Bispo Fogaca",
                "senha": "MF@7040",
                "funcao": "Supervisor(a) Eletricista Junior, I"
            },
            "marcus.souza": {
                "nome": "Marcus Vinicius Gomes de Souza",
                "senha": "MS@1035",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "mateus.silva2": {
                "nome": "Mateus Vinicius de Morais Silva",
                "senha": "MS@9987",
                "funcao": "Assistente Administrativo Pleno, V"
            },
            "meirianderson.silva": {
                "nome": "Meirianderson Caires Silva",
                "senha": "MS@5722",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "miguel.sousa": {
                "nome": "Miguel Kallebe Farias Sousa",
                "senha": "MS@4013",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "monize.fukurozaki": {
                "nome": "Monize Kozue Fukurozaki",
                "senha": "MF@9512",
                "funcao": "PS - Coordenador(a) Ambiental, PJ"
            },
            "muriel.antonio": {
                "nome": "Muriel Sousa Antonio",
                "senha": "MA@4694",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "nathalia.costa": {
                "nome": "Nathalia Marcal de Souza Costa",
                "senha": "NC@7389",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "neidilane.silva": {
                "nome": "Neidilane Costa Silva",
                "senha": "NS@8258",
                "funcao": "Supervisor(a) Administrativo Junior, I"
            },
            "nicolas.pereira": {
                "nome": "Nicolas Leite Pereira",
                "senha": "NP@9049",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "paulo.filho": {
                "nome": "Paulo Cesar Rodrigues Ribeiro Filho",
                "senha": "PF@2037",
                "funcao": "Técnico(a) Eletrotécnica Pleno, V"
            },
            "pedro.rodrigues": {
                "nome": "Pedro Henrique Pires Rodrigues",
                "senha": "PR@3852",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "rafaela.ferreira": {
                "nome": "Rafaela Corona Ferreira",
                "senha": "RF@4824",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "railson.jesus": {
                "nome": "Railson Pereira de Jesus",
                "senha": "RJ@5530",
                "funcao": "Técnico(a) Eletrotécnica Pleno, V"
            },
            "renan.santos": {
                "nome": "Renan Farias dos Santos",
                "senha": "RS@1101",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "ricardo.silva": {
                "nome": "Ricardo Rosa Silva",
                "senha": "RS@1816",
                "funcao": "Técnico(a) Eletrotécnica Pleno, V"
            },
            "sarah.martins": {
                "nome": "Sarah Sousa Martins",
                "senha": "SM@8343",
                "funcao": "Analista Administrativo Junior, I"
            },
            "sudario.santos": {
                "nome": "Sudario Macedo dos Santos",
                "senha": "SS@1358",
                "funcao": "Técnico(a) de Meio Ambiente Junior, I"
            },
            "telvio.costa": {
                "nome": "Telvio Gonzaga da Costa",
                "senha": "TC@4596",
                "funcao": "Técnico(a) Eletrotécnica Pleno, V"
            },
            "thallytha.silva": {
                "nome": "Thallytha Monnycky Silva",
                "senha": "TS@4827",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "thiago.ruiz": {
                "nome": "Thiago Barranco Ruiz",
                "senha": "TR@3154",
                "funcao": "PS - Gerente de Processos, Geral"
            },
            "vinicius.pereira": {
                "nome": "Vinicius Leite Pereira",
                "senha": "VP@0352",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "vinicius.barbosa": {
                "nome": "Vinicius Lucas Barbosa",
                "senha": "VB@4503",
                "funcao": "Técnico(a) Eletrotécnica Pleno, V"
            },
            "vinicius.oliveira": {
                "nome": "Vinicius Silva de Oliveira",
                "senha": "VO@0723",
                "funcao": "Analista Administrativo Senior, IX"
            },
            "wanessa.pereira": {
                "nome": "Wanessa Cristina Alves Pereira",
                "senha": "WP@9583",
                "funcao": "Analista Administrativo Junior, IV"
            },
            "wartelan.bitencourt": {
                "nome": "Wartelan Ribeiro Bitencourt",
                "senha": "WB@3956",
                "funcao": "Apontador(a) de Produção Pleno, V"
            },
            "welviston.ferreira": {
                "nome": "Welviston Pereira da Silva Ferreira",
                "senha": "WF@7693",
                "funcao": "Supervisor(a) Administrativo Junior, II"
            },
            "weudes.gontijo": {
                "nome": "Weudes Simoes Gontijo",
                "senha": "WG@5049",
                "funcao": "Técnico(a) Eletrotécnica Pleno, V"
            },
            "wildes.santos": {
                "nome": "Wildes Pereira dos Santos",
                "senha": "WS@4613",
                "funcao": "Técnico(a) Eletrotécnica Pleno, V"
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
# 4. PERSISTÊNCIA EM EXCEL E HISTÓRICO DE ACESSOS (COM BLOQUEIO ANTI-SOBREPOSIÇÃO)
# ==============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_DIR = os.path.join(BASE_DIR, "database")
STATIC_DIR = os.path.join(BASE_DIR, "static")
os.makedirs(DB_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

EXCEL_PATH = os.path.join(DB_DIR, "justificativas_afaplan.xlsx")
LOCK_PATH = os.path.join(DB_DIR, "justificativas_afaplan.xlsx.lock")
ACESSOS_JSON_PATH = os.path.join(DB_DIR, "historico_acessos.json")
ACESSOS_LOCK_PATH = os.path.join(DB_DIR, "historico_acessos.json.lock")

COLUNAS_EXCEL = [
    "ID_Registro",
    "Data_Hora_Registro",
    "Projeto",
    "Funcao_Subdisciplina",
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


def registrar_acesso_historico(projeto, perfil, usuario, nome_exibicao):
    with FileLock(ACESSOS_LOCK_PATH, timeout=10):
        historico = []
        if os.path.exists(ACESSOS_JSON_PATH):
            try:
                with open(ACESSOS_JSON_PATH, "r", encoding="utf-8") as f:
                    historico = json.load(f)
            except Exception:
                historico = []

        novo_evento = {
            "data_hora": agora_brasilia().strftime("%d/%m/%Y %H:%M:%S"),
            "projeto": projeto,
            "perfil": perfil,
            "usuario": usuario,
            "nome": nome_exibicao
        }
        historico.insert(0, novo_evento)
        historico = historico[:100]

        with open(ACESSOS_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump(historico, f, ensure_ascii=False, indent=4)


def ler_historico_acessos(projeto, is_lider, usuario_login, limite=5):
    with FileLock(ACESSOS_LOCK_PATH, timeout=10):
        if not os.path.exists(ACESSOS_JSON_PATH):
            return []
        try:
            with open(ACESSOS_JSON_PATH, "r", encoding="utf-8") as f:
                historico = json.load(f)
        except Exception:
            return []

    if is_lider:
        filtrados = [h for h in historico if h.get("projeto") == projeto]
    else:
        filtrados = [
            h for h in historico
            if h.get("projeto") == projeto and h.get("usuario") == usuario_login
        ]
    return filtrados[:limite]


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

            larguras = [15, 22, 18, 32, 22, 34, 20, 18, 80]
            for i, largura in enumerate(larguras, 1):
                ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = largura

            wb.save(EXCEL_PATH)


def ler_registros_excel(projeto_atual, is_lider, nome_colaborador_restrito):
    inicializar_excel()
    registros = []
    with FileLock(LOCK_PATH, timeout=15):
        wb = load_workbook(EXCEL_PATH, data_only=True)
        ws = wb["Base_Justificativas"]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            if str(row[2]) != projeto_atual:
                continue
            # Líder vê todos os registros do projeto; colaborador vê apenas os próprios
            if not is_lider and str(row[5]) != nome_colaborador_restrito:
                continue

            registros.append({
                "id": str(row[0]),
                "data_hora": str(row[1]),
                "projeto": str(row[2]),
                "funcao": str(row[3]),
                "usuario": str(row[4]),
                "colaborador": str(row[5]),
                "mes": str(row[6]),
                "semana": str(row[7]),
                "justificativa": str(row[8])
            })
        wb.close()
    registros.reverse()
    return registros


def salvar_registro_excel(projeto, funcao, usuario, colaborador, mes_op, semana_op, justificativa):
    inicializar_excel()
    with FileLock(LOCK_PATH, timeout=20):
        wb = load_workbook(EXCEL_PATH)
        ws = wb["Base_Justificativas"]
        novo_id = agora_brasilia().strftime("%Y%m%d%H%M%S%f")[:18]
        data_hora_fmt = agora_brasilia().strftime("%d/%m/%Y %H:%M:%S")
        ws.append([
            novo_id,
            data_hora_fmt,
            projeto,
            funcao,
            usuario,
            colaborador,
            mes_op,
            semana_op,
            justificativa
        ])
        wb.save(EXCEL_PATH)
        wb.close()


def excluir_registro_excel(registro_id, projeto_atual, is_lider, nome_colaborador_restrito):
    with FileLock(LOCK_PATH, timeout=20):
        wb = load_workbook(EXCEL_PATH)
        ws = wb["Base_Justificativas"]
        linha_del = None
        for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if str(row[0]) == str(registro_id) and str(row[2]) == projeto_atual:
                if is_lider or str(row[5]) == nome_colaborador_restrito:
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

# ==============================================================================
# 5. CONTROLE DE SESSÃO (LOGIN / LOGOUT)
# ==============================================================================
DEFAULTS_SESSION = {
    "logado": False,
    "projeto": "",
    "is_lider": False,
    "usuario": "",
    "nome_colaborador": "",
    "funcao_colaborador": "",
    "uf": "PA"
}
for _k, _v in DEFAULTS_SESSION.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


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
        projeto_sel = st.selectbox("1. Selecione o Projeto", list(ACESSOS_PROJETOS.keys()))

        with st.form("form_login_st"):
            usuario_inp = st.text_input("2. Login de Acesso", placeholder="Ex.: nome.sobrenome ou lider.pa")
            senha_inp = st.text_input("3. Senha", type="password", placeholder="Digite sua senha de acesso")
            entrar = st.form_submit_button("Acessar Painel", type="primary", use_container_width=True)

            if entrar:
                proj_cfg = ACESSOS_PROJETOS[projeto_sel]
                u_clean = usuario_inp.strip().lower()

                # Verifica login de Líder
                if u_clean == proj_cfg["lider"]["login"].lower() and senha_inp == proj_cfg["lider"]["senha"]:
                    st.session_state.logado = True
                    st.session_state.projeto = projeto_sel
                    st.session_state.is_lider = True
                    st.session_state.usuario = proj_cfg["lider"]["login"]
                    st.session_state.nome_colaborador = f"Líder {projeto_sel}"
                    st.session_state.funcao_colaborador = "Liderança / Gestão"
                    st.session_state.uf = proj_cfg["uf"]
                    registrar_acesso_historico(projeto_sel, "Líder", proj_cfg["lider"]["login"], f"Líder {projeto_sel}")
                    st.rerun()

                # Verifica login individual do Colaborador vinculado ao projeto
                usr_info = proj_cfg["usuarios"].get(u_clean)
                if usr_info and senha_inp == usr_info["senha"]:
                    st.session_state.logado = True
                    st.session_state.projeto = projeto_sel
                    st.session_state.is_lider = False
                    st.session_state.usuario = u_clean
                    st.session_state.nome_colaborador = usr_info["nome"]
                    st.session_state.funcao_colaborador = usr_info["funcao"]
                    st.session_state.uf = proj_cfg["uf"]
                    registrar_acesso_historico(projeto_sel, "Colaborador", u_clean, usr_info["nome"])
                    st.rerun()

                st.error("Projeto, login ou senha incorretos para este projeto.")


def tela_registro():
    projeto = st.session_state.get("projeto", "")
    is_lider = st.session_state.get("is_lider", False)
    usuario = st.session_state.get("usuario", "")
    nome_colab_logado = st.session_state.get("nome_colaborador", "")
    funcao_colab_logado = st.session_state.get("funcao_colaborador", "")
    uf_projeto = st.session_state.get("uf", "PA")

    if not projeto or projeto not in ACESSOS_PROJETOS:
        st.session_state.logado = False
        st.rerun()

    perfil_label = "Líder" if is_lider else nome_colab_logado
    render_banner(f"{projeto} • {perfil_label} ({usuario})")

    _, col_sair = st.columns([8, 1])
    with col_sair:
        if st.button("Sair", use_container_width=True):
            st.session_state.logado = False
            st.rerun()

    # Linha 1: Colaborador | Mês/Ano | Semana (Quinta a Quarta)
    idx_nov = OPCOES_MES_ANO.index("Novembro / 2026") if "Novembro / 2026" in OPCOES_MES_ANO else 0

    c1, c2, c3 = st.columns([1.5, 1.1, 1.4])
    with c1:
        if is_lider:
            todos_colabs = [u["nome"] for u in ACESSOS_PROJETOS[projeto]["usuarios"].values()]
            colaborador_sel = st.selectbox("Nome do Colaborador (Visão Líder)", todos_colabs)
            funcao_sel = next(
                (u["funcao"] for u in ACESSOS_PROJETOS[projeto]["usuarios"].values() if u["nome"] == colaborador_sel),
                ""
            )
        else:
            # Colaborador comum pode selecionar ESTRITAMENTE apenas o próprio nome
            colaborador_sel = st.selectbox(
                "Nome do Colaborador",
                [nome_colab_logado],
                index=0,
                disabled=True
            )
            funcao_sel = funcao_colab_logado

    with c2:
        mes_sel = st.selectbox("Mês/Ano de Operação", OPCOES_MES_ANO, index=idx_nov)
    with c3:
        semanas_mes = gerar_semanas_quinta_quarta(mes_sel)
        rotulos_semanas = [s["rotulo"] for s in semanas_mes]
        semana_rotulo_sel = st.selectbox("Semana de Operação (Quinta a Quarta)", rotulos_semanas)

    semana_obj = next(s for s in semanas_mes if s["rotulo"] == semana_rotulo_sel)
    semana_chave = semana_obj["chave"]

    # Linha 2: Justificativa Diária (Esquerda) + Notas de Feriados e Histórico de Acessos (Direita)
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
            key_input = f"inp_{usuario}_{mes_sel}_{semana_chave}_{data_curta}"

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
                    else f"Ex.: Fiquei doente / Atestado médico / Não se aplica ({data_curta})"
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
                f'<div style="color:#486581; font-style:italic; font-size:0.88rem;">'
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

        # Histórico de Acessos dentro do retângulo inferior direito
        acessos = ler_historico_acessos(projeto, is_lider, usuario, limite=5)
        if not acessos:
            acessos_html = (
                '<div class="acessos-card">'
                '<div class="acessos-titulo">Histórico de Acessos</div>'
                '<div class="acesso-vazio">Nenhum acesso registrado até o momento.</div>'
                '</div>'
            )
        else:
            ac_parts = []
            for ac in acessos:
                ac_parts.append(
                    f'<div class="acesso-item">'
                    f'<div class="acesso-hora">{ac.get("data_hora", "")}</div>'
                    f'<div><strong>{ac.get("nome", "")}</strong> ({ac.get("usuario", "")})</div>'
                    f'<div>{ac.get("projeto", "")} • {ac.get("perfil", "")}</div>'
                    f'</div>'
                )
            acessos_html = (
                '<div class="acessos-card">'
                '<div class="acessos-titulo">Histórico de Acessos</div>'
                + "".join(ac_parts)
                + '</div>'
            )
        st.markdown(acessos_html, unsafe_allow_html=True)

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
                k = f"inp_{usuario}_{mes_sel}_{semana_chave}_{d.strftime('%d/%m')}"
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
                    funcao=funcao_sel,
                    usuario=usuario,
                    colaborador=colaborador_sel,
                    mes_op=mes_sel,
                    semana_op=semana_chave,
                    justificativa=justificativa_somada
                )
                for d in semana_obj["dias"]:
                    k = f"inp_{usuario}_{mes_sel}_{semana_chave}_{d.strftime('%d/%m')}"
                    if k in st.session_state:
                        st.session_state[k] = ""
                st.success("Justificativa semanal consolidada e registrada no Excel com sucesso!")
                st.rerun()

    # Histórico de Justificativas (Botão de Download habilitado EXCLUSIVAMENTE para o Líder)
    st.markdown('<div class="hist-card">', unsafe_allow_html=True)
    if is_lider:
        col_hist_tit, col_dl = st.columns([4, 1.4])
        with col_hist_tit:
            st.markdown("### Histórico de Justificativas (Visão Geral do Líder)")
        with col_dl:
            if os.path.exists(EXCEL_PATH):
                with open(EXCEL_PATH, "rb") as f_xlsx:
                    st.download_button(
                        label="📥 Baixar Excel (.xlsx)",
                        data=f_xlsx.read(),
                        file_name="justificativas_afaplan.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
    else:
        st.markdown("### Histórico de Justificativas")

    registros = ler_registros_excel(projeto, is_lider, nome_colab_logado)
    if not registros:
        st.markdown(
            '<div class="hist-vazio-box">Nenhuma justificativa registrada até o momento.</div>',
            unsafe_allow_html=True
        )
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
                excluir_registro_excel(reg["id"], projeto, is_lider, nome_colab_logado)
                st.rerun()
            st.markdown("<hr style='margin: 6px 0; border-top: 1px solid #e4e7eb;'>", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)


if st.session_state.logado:
    tela_registro()
else:
    tela_login()
