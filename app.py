import os
import json
import calendar
from datetime import datetime, date, timedelta
from functools import wraps
from flask import Flask, request, jsonify, render_template_string, session, redirect, url_for
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from filelock import FileLock

app = Flask(__name__)
app.secret_key = "afaplan_chave_super_secreta_2026_v5"

# ==============================================================================
# 1. ESTRUTURA DE PROJETOS, SUBDISCIPLINAS, CREDENCIAIS E COLABORADORES
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
# 2. FERIADOS NACIONAIS E REGIONAIS (PARÁ, AMAPÁ E GOIÁS) + CICLO QUINTA A QUARTA
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
    """
    Calcula as semanas de operação do mês/ano onde cada semana cheia inicia
    na Quinta-feira (weekday == 3) e fecha na Quarta-feira (weekday == 2).
    """
    partes = [p.strip() for p in mes_ano_str.split("/")]
    nome_mes = partes[0]
    ano = int(partes[1])
    mes_num = {v: k for k, v in MESES_PT.items()}[nome_mes]

    primeiro_dia = date(ano, mes_num, 1)
    ultimo_dia = date(ano, mes_num, calendar.monthrange(ano, mes_num)[1])

    semanas = []
    # Encontra a primeira quinta-feira do mês ou a última quinta-feira do mês anterior
    # para que a contagem da semana cheia inicie na quinta-feira e feche na quarta-feira.
    offset_para_quinta = (primeiro_dia.weekday() - 3) % 7
    inicio_ciclo = primeiro_dia - timedelta(days=offset_para_quinta)

    # Se o início do ciclo caiu no mês anterior e tem menos de 4 dias no mês atual,
    # inicia na primeira quinta-feira dentro do próprio mês; caso contrário inclui o ciclo.
    dias_no_mes = (inicio_ciclo + timedelta(days=6) - primeiro_dia).days + 1
    if inicio_ciclo < primeiro_dia and dias_no_mes < 4:
        inicio_ciclo = inicio_ciclo + timedelta(days=7)

    idx = 1
    atual_ini = inicio_ciclo
    while atual_ini <= ultimo_dia:
        atual_fim = atual_ini + timedelta(days=6)  # Quinta (+6 dias) = Quarta-feira
        dias_lista = []
        for i in range(7):
            d = atual_ini + timedelta(days=i)
            dias_lista.append(d)

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


def consultar_dados_semana(uf_projeto, mes_ano_str, semana_chave):
    """
    Retorna os dias da semana selecionada (Quinta a Quarta, incluindo feriados)
    e a lista de feriados nacionais + estaduais (PA/AP/GO) do período.
    """
    semanas = gerar_semanas_quinta_quarta(mes_ano_str)
    semana_obj = next((s for s in semanas if s["chave"] == semana_chave), semanas[0])

    dias_detalhados = []
    feriados_encontrados = []

    for d in semana_obj["dias"]:
        feriado_info = None
        info = FERIADOS_CADASTRO.get((d.month, d.day))
        if info and info["escopo"] in ("Nacional", uf_projeto):
            tipo_label = "Feriado Nacional" if info["escopo"] == "Nacional" else f"Feriado Regional ({info['escopo']})"
            feriado_info = {
                "data": d.strftime("%d/%m/%Y"),
                "dia_extenso": f"{d.day:02d} de {MESES_PT[d.month].lower()}",
                "nome": info["nome"],
                "tipo": tipo_label
            }
            feriados_encontrados.append(feriado_info)

        dias_detalhados.append({
            "data_curta": d.strftime("%d/%m"),
            "data_completa": d.strftime("%d/%m/%Y"),
            "dia_semana": DIAS_SEMANA_PT[d.weekday()],
            "is_fim_semana": d.weekday() in (5, 6),
            "feriado": feriado_info
        })

    opcoes_semanas = [{"chave": s["chave"], "rotulo": s["rotulo"]} for s in semanas]

    return {
        "opcoes_semanas": opcoes_semanas,
        "semana_selecionada": semana_obj["chave"],
        "periodo_texto": f"{semana_obj['inicio'].strftime('%d/%m/%Y')} (Quinta) a {semana_obj['fim'].strftime('%d/%m/%Y')} (Quarta)",
        "dias": dias_detalhados,
        "feriados": feriados_encontrados
    }


# ==============================================================================
# 3. PERSISTÊNCIA EM JSON E EXCEL (COM BLOQUEIO ANTI-SOBREPOSIÇÃO)
# ==============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_DIR = os.path.join(BASE_DIR, "database")
STATIC_DIR = os.path.join(BASE_DIR, "static")
os.makedirs(DB_DIR, exist_ok=True)
os.makedirs(STATIC_DIR, exist_ok=True)

GRUPOS_JSON_PATH = os.path.join(DB_DIR, "colaboradores_por_grupo.json")
EXCEL_PATH = os.path.join(DB_DIR, "justificativas_afaplan.xlsx")
LOCK_PATH = os.path.join(DB_DIR, "justificativas_afaplan.xlsx.lock")


def localizar_logo():
    candidatos = [
        "Logo_afaplan_1.png",
        "Logo_afaplan_1 (1).png",
        "logo_afaplan_1.png",
        "Logo_afaplan_2.png",
        "Logo_afaplan_2 (1).png"
    ]
    for nome in candidatos:
        if os.path.exists(os.path.join(STATIC_DIR, nome)):
            return f"/static/{nome}"

    for arq in sorted(os.listdir(STATIC_DIR)):
        if arq.lower().endswith((".png", ".jpg", ".jpeg", ".webp")):
            return f"/static/{arq}"

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


inicializar_excel()
carregar_estrutura_grupos()


def login_requerido(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "projeto" not in session or "subdisciplina" not in session or "usuario" not in session:
            return redirect(url_for("login_page"))
        return f(*args, **kwargs)
    return decorated_function


# ==============================================================================
# 4. TEMPLATES HTML/CSS/JS
# ==============================================================================
LOGIN_HTML = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Afaplan - Login de Justificativas</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; }
        body {
            background: linear-gradient(135deg, #f0f4f8 0%, #d9e2ec 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        .login-card {
            background: #ffffff;
            width: 100%;
            max-width: 440px;
            border-radius: 12px;
            box-shadow: 0 10px 30px rgba(0, 0, 0, 0.12);
            overflow: hidden;
        }
        .login-header {
            background: #ffffff;
            padding: 28px 25px 20px;
            text-align: center;
            border-bottom: 4px solid #1560bd;
        }
        .login-header img {
            max-height: 52px;
            max-width: 220px;
            width: auto;
            height: auto;
            margin-bottom: 12px;
            display: inline-block;
        }
        .login-header h2 {
            color: #1f2933;
            font-size: 1.12rem;
            font-weight: 700;
            letter-spacing: 0.5px;
        }
        .login-body {
            padding: 26px 30px 30px;
        }
        .form-group {
            margin-bottom: 16px;
        }
        label {
            display: block;
            margin-bottom: 6px;
            font-weight: 600;
            color: #334e68;
            font-size: 0.92rem;
        }
        select, input {
            width: 100%;
            padding: 11px 14px;
            border: 1px solid #cbd2d9;
            border-radius: 6px;
            font-size: 0.96rem;
            color: #102a43;
            background-color: #fff;
        }
        select:focus, input:focus {
            outline: none;
            border-color: #1560bd;
            box-shadow: 0 0 0 3px rgba(21, 96, 189, 0.18);
        }
        .btn-login {
            width: 100%;
            padding: 13px;
            background-color: #1560bd;
            color: #fff;
            border: none;
            border-radius: 6px;
            font-size: 1rem;
            font-weight: 600;
            cursor: pointer;
            margin-top: 8px;
        }
        .btn-login:hover {
            background-color: #0f4c9a;
        }
        .erro-msg {
            background: #ffe3e3;
            color: #c92a2a;
            padding: 10px 12px;
            border-radius: 6px;
            font-size: 0.88rem;
            margin-bottom: 16px;
            display: none;
            border-left: 4px solid #c92a2a;
        }
    </style>
</head>
<body>
    <div class="login-card">
        <div class="login-header">
            {% if logo_url %}
                <img src="{{ logo_url }}" alt="Afaplan">
            {% else %}
                <svg width="170" height="44" viewBox="0 0 170 44" xmlns="http://www.w3.org/2000/svg" style="margin-bottom:10px;">
                    <path d="M8 36 C8 16, 22 8, 40 8 L40 36 Z" fill="none" stroke="#1560bd" stroke-width="6"/>
                    <text x="50" y="30" font-family="Segoe UI, Arial, sans-serif" font-size="26" font-weight="700" fill="#2d3136">afaplan</text>
                </svg>
            {% endif %}
            <h2>SISTEMA DE JUSTIFICATIVAS</h2>
        </div>
        <div class="login-body">
            <div id="erroBox" class="erro-msg"></div>
            <form id="loginForm">
                <div class="form-group">
                    <label for="projeto">1. Selecione o Projeto</label>
                    <select id="projeto" required onchange="atualizarSubdisciplinas()">
                        <option value="" disabled selected>Escolha o projeto...</option>
                        {% for proj in projetos %}
                        <option value="{{ proj }}">{{ proj }}</option>
                        {% endfor %}
                    </select>
                </div>

                <div class="form-group">
                    <label for="subdisciplina">2. Selecione a Subdisciplina</label>
                    <select id="subdisciplina" required disabled>
                        <option value="" disabled selected>Selecione o projeto primeiro...</option>
                    </select>
                </div>

                <div class="form-group">
                    <label for="login">3. Login de Acesso</label>
                    <input type="text" id="login" placeholder="Digite seu usuário" required autocomplete="username">
                </div>

                <div class="form-group">
                    <label for="senha">4. Senha</label>
                    <input type="password" id="senha" placeholder="Digite sua senha" required autocomplete="current-password">
                </div>

                <button type="submit" class="btn-login">Acessar Painel</button>
            </form>
        </div>
    </div>

    <script>
        const mapaSubdisciplinas = {{ mapa_subdisciplinas_json | safe }};

        function atualizarSubdisciplinas() {
            const projeto = document.getElementById("projeto").value;
            const selectSub = document.getElementById("subdisciplina");
            selectSub.innerHTML = '<option value="" disabled selected>Escolha a subdisciplina...</option>';

            if (projeto && mapaSubdisciplinas[projeto]) {
                mapaSubdisciplinas[projeto].forEach(sub => {
                    const opt = document.createElement("option");
                    opt.value = sub;
                    opt.textContent = sub;
                    selectSub.appendChild(opt);
                });
                selectSub.disabled = false;
            } else {
                selectSub.disabled = true;
            }
        }

        document.getElementById("loginForm").addEventListener("submit", async (e) => {
            e.preventDefault();
            const erroBox = document.getElementById("erroBox");
            erroBox.style.display = "none";

            const payload = {
                projeto: document.getElementById("projeto").value,
                subdisciplina: document.getElementById("subdisciplina").value,
                login: document.getElementById("login").value.trim(),
                senha: document.getElementById("senha").value
            };

            const res = await fetch("/api/login", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload)
            });
            const data = await res.json();

            if (data.sucesso) {
                window.location.href = "/registro";
            } else {
                erroBox.textContent = data.mensagem || "Credenciais inválidas.";
                erroBox.style.display = "block";
            }
        });
    </script>
</body>
</html>
"""

PAINEL_HTML = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Registro de Justificativas de Desempenho - Afaplan</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; }
        body {
            background-color: #f4f6f9;
            color: #1f2933;
            min-height: 100vh;
        }

        .top-banner {
            display: flex;
            align-items: stretch;
            background-color: #ffffff;
            box-shadow: 0 2px 6px rgba(0,0,0,0.08);
            min-height: 76px;
        }
        .logo-area {
            background-color: #ffffff;
            padding: 12px 24px;
            display: flex;
            align-items: center;
            justify-content: center;
            min-width: 220px;
        }
        .logo-area img {
            height: 44px;
            width: auto;
        }
        .title-area {
            flex: 1;
            background-color: #1560bd;
            color: #ffffff;
            display: flex;
            align-items: center;
            justify-content: space-between;
            padding: 12px 28px;
            gap: 15px;
        }
        .title-area h1 {
            font-size: 1.35rem;
            font-weight: 700;
            letter-spacing: 0.5px;
        }
        .user-badge {
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 0.85rem;
            background: rgba(255,255,255,0.16);
            padding: 6px 14px;
            border-radius: 20px;
            white-space: nowrap;
        }
        .btn-sair {
            background: #ffffff;
            color: #1560bd;
            border: none;
            padding: 4px 12px;
            border-radius: 14px;
            font-weight: 700;
            cursor: pointer;
            font-size: 0.8rem;
            text-decoration: none;
        }
        .btn-sair:hover { background: #e6f0fa; }

        .container {
            max-width: 1180px;
            margin: 28px auto;
            padding: 0 20px 50px;
        }

        .selectors-grid {
            display: grid;
            grid-template-columns: 1.6fr 1.1fr 1.3fr;
            gap: 18px;
            margin-bottom: 22px;
        }
        .field-group label {
            display: block;
            font-size: 0.95rem;
            font-weight: 600;
            color: #1f2933;
            margin-bottom: 7px;
        }
        .field-group select {
            width: 100%;
            background-color: #ffffff;
            border: 1px solid #cbd2d9;
            border-radius: 6px;
            padding: 11px 14px;
            font-size: 0.96rem;
            color: #102a43;
            box-shadow: 0 1px 2px rgba(0,0,0,0.04);
        }
        .field-group select:focus {
            outline: none;
            border-color: #1560bd;
            box-shadow: 0 0 0 3px rgba(21, 96, 189, 0.15);
        }

        /* LINHA DE JUSTIFICATIVA DIÁRIA (ESQUERDA) + NOTAS DE FERIADOS (DIREITA) */
        .desc-notas-grid {
            display: grid;
            grid-template-columns: 1.75fr 1fr;
            gap: 20px;
            margin-top: 6px;
            margin-bottom: 28px;
            align-items: start;
        }

        .dias-card {
            background: #ffffff;
            border: 1px solid #cbd2d9;
            border-radius: 8px;
            padding: 14px 16px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        }
        .dias-header-info {
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 8px;
            padding-bottom: 10px;
            margin-bottom: 12px;
            border-bottom: 1px solid #e4e7eb;
            font-size: 0.85rem;
            color: #486581;
        }
        .dias-lista {
            display: flex;
            flex-direction: column;
            gap: 10px;
        }
        .dia-row {
            display: grid;
            grid-template-columns: 185px 1fr;
            gap: 12px;
            align-items: center;
            background: #f8fafc;
            border: 1px solid #e4e7eb;
            border-radius: 6px;
            padding: 8px 10px;
        }
        .dia-row.is-feriado {
            background: #fff9db;
            border-color: #f59f00;
        }
        .dia-row.is-fds {
            background: #f1f5f9;
        }
        .dia-label-box {
            display: flex;
            flex-direction: column;
            gap: 2px;
        }
        .dia-data {
            font-weight: 700;
            font-size: 0.92rem;
            color: #102a43;
        }
        .dia-nome {
            font-size: 0.80rem;
            color: #486581;
        }
        .dia-tag-feriado {
            display: inline-block;
            font-size: 0.72rem;
            font-weight: 700;
            color: #b45309;
            background: #fef3c7;
            border: 1px solid #fcd34d;
            border-radius: 4px;
            padding: 1px 6px;
            margin-top: 2px;
            width: fit-content;
        }
        .dia-input {
            width: 100%;
            background: #ffffff;
            border: 1px solid #cbd2d9;
            border-radius: 5px;
            padding: 9px 12px;
            font-size: 0.92rem;
            color: #102a43;
        }
        .dia-input:focus {
            outline: none;
            border-color: #1560bd;
            box-shadow: 0 0 0 3px rgba(21, 96, 189, 0.15);
        }

        .notas-box {
            background-color: #fffdf5;
            border: 1px solid #f0d58c;
            border-left: 5px solid #e67700;
            border-radius: 8px;
            padding: 14px 16px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            display: flex;
            flex-direction: column;
            justify-content: flex-start;
        }
        .notas-header {
            font-size: 0.94rem;
            font-weight: 700;
            color: #9c4221;
            margin-bottom: 10px;
            display: flex;
            align-items: center;
            justify-content: space-between;
        }
        .notas-content {
            font-size: 0.88rem;
            color: #334e68;
            line-height: 1.45;
        }
        .feriado-item {
            background: #ffffff;
            border: 1px solid #fae3a8;
            border-radius: 5px;
            padding: 8px 10px;
            margin-bottom: 8px;
        }
        .feriado-item:last-child { margin-bottom: 0; }
        .feriado-badge {
            display: inline-block;
            font-size: 0.74rem;
            font-weight: 700;
            padding: 2px 7px;
            border-radius: 10px;
            background: #e7f5ff;
            color: #1560bd;
            margin-left: 6px;
        }
        .feriado-badge.regional {
            background: #fff4e6;
            color: #d9480f;
        }
        .sem-feriado {
            color: #627d98;
            font-style: italic;
            font-size: 0.86rem;
        }

        .actions-bar {
            position: relative;
            z-index: 2;
            clear: both;
            display: flex;
            justify-content: center;
            gap: 18px;
            margin-top: 10px;
            margin-bottom: 32px;
            flex-wrap: wrap;
        }
        .btn {
            min-width: 145px;
            padding: 11px 24px;
            border-radius: 6px;
            font-size: 0.98rem;
            font-weight: 500;
            cursor: pointer;
            transition: all 0.2s;
            box-shadow: 0 2px 4px rgba(0,0,0,0.08);
        }
        .btn-cancelar {
            background-color: #ffffff;
            color: #c92a2a;
            border: 1px solid #e03131;
        }
        .btn-cancelar:hover { background-color: #fff5f5; }
        .btn-limpar {
            background-color: #3e4c59;
            color: #ffffff;
            border: 1px solid #3e4c59;
        }
        .btn-limpar:hover { background-color: #2c3742; }
        .btn-cadastrar {
            background-color: #1560bd;
            color: #ffffff;
            border: 1px solid #1560bd;
        }
        .btn-cadastrar:hover { background-color: #0f4c9a; }

        .table-card {
            background: #ffffff;
            border-radius: 8px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.07);
            overflow-x: auto;
            border: 1px solid #d9e2ec;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            min-width: 850px;
        }
        thead {
            background-color: #f8fafc;
            border-bottom: 2px solid #d9e2ec;
        }
        th {
            text-align: left;
            padding: 14px 16px;
            font-size: 0.92rem;
            font-weight: 700;
            color: #102a43;
        }
        td {
            padding: 14px 16px;
            font-size: 0.92rem;
            color: #334e68;
            border-bottom: 1px solid #e4e7eb;
            vertical-align: top;
            line-height: 1.45;
        }
        tr:hover td { background-color: #f9fbfd; }
        .btn-del {
            background: #ffe3e3;
            color: #c92a2a;
            border: 1px solid #ffa8a8;
            padding: 5px 10px;
            border-radius: 4px;
            cursor: pointer;
            font-size: 0.82rem;
            font-weight: 600;
        }
        .btn-del:hover { background: #ffc9c9; }

        .toast {
            position: fixed;
            bottom: 22px;
            right: 22px;
            background: #1560bd;
            color: #fff;
            padding: 12px 20px;
            border-radius: 6px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.2);
            display: none;
            z-index: 999;
            font-size: 0.92rem;
        }

        @media (max-width: 900px) {
            .top-banner { flex-direction: column; }
            .logo-area { padding: 12px; border-bottom: 1px solid #e4e7eb; }
            .title-area { flex-direction: column; text-align: center; padding: 15px; }
            .title-area h1 { font-size: 1.12rem; }
            .selectors-grid { grid-template-columns: 1fr; gap: 14px; }
            .desc-notas-grid { grid-template-columns: 1fr; gap: 14px; }
            .dia-row { grid-template-columns: 1fr; gap: 6px; }
            .btn { width: 100%; }
        }
    </style>
</head>
<body>

    <header class="top-banner">
        <div class="logo-area">
            {% if logo_url %}
                <img src="{{ logo_url }}" alt="Afaplan">
            {% else %}
                <svg width="160" height="42" viewBox="0 0 170 44" xmlns="http://www.w3.org/2000/svg">
                    <path d="M8 36 C8 16, 22 8, 40 8 L40 36 Z" fill="none" stroke="#1560bd" stroke-width="6"/>
                    <text x="50" y="30" font-family="Segoe UI, Arial, sans-serif" font-size="26" font-weight="700" fill="#2d3136">afaplan</text>
                </svg>
            {% endif %}
        </div>
        <div class="title-area">
            <h1>REGISTRO DE JUSTIFICATIVAS DE DESEMPENHO</h1>
            <div class="user-badge">
                <span><strong>{{ projeto }}</strong> • <strong>{{ subdisciplina }}</strong> ({{ usuario }})</span>
                <a href="/logout" class="btn-sair">Sair</a>
            </div>
        </div>
    </header>

    <main class="container">
        <form id="registroForm">
            <div class="selectors-grid">
                <div class="field-group">
                    <label for="colaborador">Nome do Colaborador</label>
                    <select id="colaborador" required>
                        {% for colab in colaboradores %}
                        <option value="{{ colab }}">{{ colab }}</option>
                        {% endfor %}
                    </select>
                </div>

                <div class="field-group">
                    <label for="mesOperacao">Mês/Ano de Operação</label>
                    <select id="mesOperacao" required onchange="atualizarSemanaEDias(true)">
                        {% for mes_item in opcoes_mes_ano %}
                        <option value="{{ mes_item }}" {% if mes_item == 'Novembro / 2026' %}selected{% endif %}>{{ mes_item }}</option>
                        {% endfor %}
                    </select>
                </div>

                <div class="field-group">
                    <label for="semanaOperacao">Semana de Operação (Quinta a Quarta)</label>
                    <select id="semanaOperacao" required onchange="atualizarSemanaEDias(false)">
                        <option value="Semana 1">Semana 1</option>
                    </select>
                </div>
            </div>

            <div class="desc-notas-grid">
                <div class="field-group">
                    <label>Descrição Diária dos Acontecimentos da Semana (Justificativa por dia)</label>
                    <div class="dias-card">
                        <div class="dias-header-info">
                            <span><strong>Período da Semana:</strong> <span id="periodoSemanaTexto">Carregando...</span></span>
                            <span>Ciclo: Quinta-feira a Quarta-feira</span>
                        </div>
                        <div id="listaDiasSemana" class="dias-lista">
                            Carregando dias da semana...
                        </div>
                    </div>
                </div>

                <div class="field-group">
                    <label for="notasFeriados">Notas</label>
                    <div id="notasFeriados" class="notas-box">
                        <div class="notas-header">
                            <span>Feriados da Semana ({{ uf_projeto }})</span>
                        </div>
                        <div id="notasConteudo" class="notas-content">Carregando...</div>
                    </div>
                </div>
            </div>

            <div class="actions-bar">
                <button type="button" class="btn btn-cancelar" onclick="window.location.href='/logout'">Cancelar</button>
                <button type="button" class="btn btn-limpar" onclick="limparFormulario()">Limpar</button>
                <button type="submit" id="btnCadastrar" class="btn btn-cadastrar">Cadastrar</button>
            </div>
        </form>

        <div class="table-card">
            <table>
                <thead>
                    <tr>
                        <th style="width: 165px;">Data/Hora</th>
                        <th style="width: 175px;">Colaborador</th>
                        <th style="width: 145px;">Mês de Operação</th>
                        <th style="width: 105px;">Semana</th>
                        <th>Descrição da Semana</th>
                        <th style="width: 90px; text-align: center;">Ações</th>
                    </tr>
                </thead>
                <tbody id="tabelaRegistros">
                    <tr><td colspan="6" style="text-align:center;">Carregando registros...</td></tr>
                </tbody>
            </table>
        </div>
    </main>

    <div id="toast" class="toast"></div>

    <script>
        function mostrarToast(msg, erro = false) {
            const toast = document.getElementById("toast");
            toast.textContent = msg;
            toast.style.backgroundColor = erro ? "#c92a2a" : "#1560bd";
            toast.style.display = "block";
            setTimeout(() => { toast.style.display = "none"; }, 3500);
        }

        function limparFormulario() {
            document.querySelectorAll(".dia-input").forEach(inp => {
                inp.value = "";
            });
        }

        async function atualizarSemanaEDias(resetSemana = false) {
            const mes = document.getElementById("mesOperacao").value;
            const selectSemana = document.getElementById("semanaOperacao");
            const semanaAtual = resetSemana ? "Semana 1" : (selectSemana.value || "Semana 1");

            const boxNotas = document.getElementById("notasConteudo");
            const listaDias = document.getElementById("listaDiasSemana");
            const periodoTxt = document.getElementById("periodoSemanaTexto");

            try {
                const params = new URLSearchParams({ mes, semana: semanaAtual });
                const res = await fetch(`/api/semana_info?${params.toString()}`);
                const data = await res.json();

                // Atualiza dropdown de semanas (Quinta a Quarta)
                selectSemana.innerHTML = "";
                data.opcoes_semanas.forEach(op => {
                    const opt = document.createElement("option");
                    opt.value = op.chave;
                    opt.textContent = op.rotulo;
                    if (op.chave === data.semana_selecionada) opt.selected = true;
                    selectSemana.appendChild(opt);
                });

                periodoTxt.textContent = data.periodo_texto;

                // Renderiza os 7 dias da semana (Quinta a Quarta), inclusive feriados
                listaDias.innerHTML = data.dias.map(d => {
                    const rowClass = d.feriado ? "dia-row is-feriado" : (d.is_fim_semana ? "dia-row is-fds" : "dia-row");
                    const feriadoBadge = d.feriado
                        ? `<span class="dia-tag-feriado">Feriado: ${d.feriado.nome}</span>`
                        : "";
                    const placeholderTxt = d.feriado
                        ? `Ex.: Feriado (${d.feriado.nome}) ou ocorrência do dia ${d.data_curta}...`
                        : `Ocorrência do dia ${d.data_curta} (Ex.: Atestado médico, Não se aplica...)`;

                    return `
                        <div class="${rowClass}">
                            <div class="dia-label-box">
                                <span class="dia-data">${d.data_curta} — ${d.dia_semana}</span>
                                <span class="dia-nome">${d.data_completa}</span>
                                ${feriadoBadge}
                            </div>
                            <div>
                                <input
                                    type="text"
                                    class="dia-input"
                                    data-data-curta="${d.data_curta}"
                                    data-dia-semana="${d.dia_semana}"
                                    placeholder="${placeholderTxt}"
                                >
                            </div>
                        </div>
                    `;
                }).join("");

                // Atualiza painel de Notas (Feriados Nacionais + Regionais PA/AP/GO)
                if (!data.feriados || data.feriados.length === 0) {
                    boxNotas.innerHTML = `<span class="sem-feriado">Nenhum feriado nacional ou regional (${data.uf}) identificado para ${data.semana_selecionada.toLowerCase()} de ${mes.toLowerCase()}.</span>`;
                } else {
                    boxNotas.innerHTML = data.feriados.map(f => {
                        const isRegional = !f.tipo.includes("Nacional");
                        const badgeClass = isRegional ? "feriado-badge regional" : "feriado-badge";
                        return `
                            <div class="feriado-item">
                                <div>
                                    <strong>${f.dia_extenso} (${f.data}) - ${f.nome}</strong>
                                    <span class="${badgeClass}">${f.tipo}</span>
                                </div>
                            </div>
                        `;
                    }).join("");
                }
            } catch (e) {
                listaDias.innerHTML = `<span class="sem-feriado">Erro ao carregar os dias da semana.</span>`;
                boxNotas.innerHTML = `<span class="sem-feriado">Não foi possível consultar os feriados.</span>`;
            }
        }

        async function carregarRegistros() {
            const res = await fetch("/api/registros");
            const data = await res.json();
            const tbody = document.getElementById("tabelaRegistros");
            tbody.innerHTML = "";

            if (!data.registros || data.registros.length === 0) {
                tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; color:#627d98;">Nenhuma justificativa registrada para este projeto e subdisciplina ainda.</td></tr>';
                return;
            }

            data.registros.forEach(reg => {
                const tr = document.createElement("tr");
                tr.innerHTML = `
                    <td>${reg.data_hora}</td>
                    <td><strong>${reg.colaborador}</strong></td>
                    <td>${reg.mes}</td>
                    <td>${reg.semana}</td>
                    <td>${reg.justificativa}</td>
                    <td style="text-align:center;">
                        <button class="btn-del" onclick="excluirRegistro('${reg.id}')">Excluir</button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        document.getElementById("registroForm").addEventListener("submit", async (e) => {
            e.preventDefault();
            const btn = document.getElementById("btnCadastrar");

            // Monta a soma das justificativas diárias no formato: "23/11 - Texto; 24/11 - Texto..."
            const inputsDias = Array.from(document.querySelectorAll(".dia-input"));
            const partesPreenchidas = [];

            inputsDias.forEach(inp => {
                const texto = inp.value.trim();
                const dataCurta = inp.getAttribute("data-data-curta");
                if (texto) {
                    partesPreenchidas.push(`${dataCurta} - ${texto}`);
                }
            });

            if (partesPreenchidas.length === 0) {
                mostrarToast("Preencha a justificativa de pelo menos um dia da semana.", true);
                return;
            }

            const justificativaConsolidada = partesPreenchidas.join("; ") + ".";

            btn.disabled = true;
            btn.textContent = "Salvando...";

            const payload = {
                colaborador: document.getElementById("colaborador").value,
                mes_operacao: document.getElementById("mesOperacao").value,
                semana_operacao: document.getElementById("semanaOperacao").value,
                justificativa: justificativaConsolidada
            };

            try {
                const res = await fetch("/api/registros", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify(payload)
                });
                const data = await res.json();

                if (data.sucesso) {
                    mostrarToast("Justificativa semanal consolidada e gravada no Excel!");
                    limparFormulario();
                    await carregarRegistros();
                } else {
                    mostrarToast(data.mensagem || "Erro ao gravar.", true);
                }
            } catch (err) {
                mostrarToast("Falha de comunicação com o servidor.", true);
            } finally {
                btn.disabled = false;
                btn.textContent = "Cadastrar";
            }
        });

        async function excluirRegistro(id) {
            if (!confirm("Deseja realmente excluir este registro do banco de dados Excel?")) return;

            const res = await fetch(`/api/registros/${id}`, { method: "DELETE" });
            const data = await res.json();
            if (data.sucesso) {
                mostrarToast("Registro removido.");
                await carregarRegistros();
            } else {
                mostrarToast(data.mensagem || "Erro ao excluir.", true);
            }
        }

        atualizarSemanaEDias(true);
        carregarRegistros();
        setInterval(carregarRegistros, 15000);
    </script>
</body>
</html>
"""


# ==============================================================================
# 5. ROTAS DA APLICAÇÃO E API REST
# ==============================================================================
@app.route("/")
def login_page():
    if "projeto" in session and "subdisciplina" in session:
        return redirect(url_for("registro_page"))

    grupos = carregar_estrutura_grupos()
    projetos = list(grupos.keys())
    mapa_sub = {proj: list(cfg["subdisciplinas"].keys()) for proj, cfg in grupos.items()}

    return render_template_string(
        LOGIN_HTML,
        projetos=projetos,
        mapa_subdisciplinas_json=json.dumps(mapa_sub, ensure_ascii=False),
        logo_url=localizar_logo()
    )


@app.route("/api/login", methods=["POST"])
def api_login():
    dados = request.get_json() or {}
    projeto = dados.get("projeto")
    subdisciplina = dados.get("subdisciplina")
    login = (dados.get("login") or "").strip()
    senha = dados.get("senha") or ""

    grupos = carregar_estrutura_grupos()
    proj_cfg = grupos.get(projeto)
    if not proj_cfg:
        return jsonify({"sucesso": False, "mensagem": "Projeto inválido."}), 401

    sub_cfg = proj_cfg.get("subdisciplinas", {}).get(subdisciplina)
    if not sub_cfg:
        return jsonify({"sucesso": False, "mensagem": "Subdisciplina inválida para o projeto selecionado."}), 401

    cred_sub_valida = (login == sub_cfg.get("login") and senha == sub_cfg.get("senha"))
    cred_proj_valida = (login == proj_cfg.get("login_geral") and senha == proj_cfg.get("senha_geral"))

    if cred_sub_valida or cred_proj_valida:
        session["projeto"] = projeto
        session["subdisciplina"] = subdisciplina
        session["usuario"] = login
        session["uf"] = proj_cfg.get("uf", "PA")
        return jsonify({"sucesso": True})

    return jsonify({"sucesso": False, "mensagem": "Projeto, subdisciplina, login ou senha incorretos."}), 401


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login_page"))


@app.route("/registro")
@login_requerido
def registro_page():
    projeto = session["projeto"]
    subdisciplina = session["subdisciplina"]
    usuario = session["usuario"]
    uf_projeto = session.get("uf", "PA")

    grupos = carregar_estrutura_grupos()
    colaboradores = grupos[projeto]["subdisciplinas"][subdisciplina]["colaboradores"]

    return render_template_string(
        PAINEL_HTML,
        projeto=projeto,
        subdisciplina=subdisciplina,
        usuario=usuario,
        uf_projeto=uf_projeto,
        colaboradores=colaboradores,
        opcoes_mes_ano=OPCOES_MES_ANO,
        logo_url=localizar_logo()
    )


@app.route("/api/semana_info", methods=["GET"])
@login_requerido
def api_semana_info():
    mes = request.args.get("mes", "Novembro / 2026")
    semana = request.args.get("semana", "Semana 1")
    uf = session.get("uf", "PA")
    dados_semana = consultar_dados_semana(uf, mes, semana)
    return jsonify({"sucesso": True, "uf": uf, **dados_semana})


@app.route("/api/registros", methods=["GET"])
@login_requerido
def listar_registros():
    projeto_atual = session["projeto"]
    sub_atual = session["subdisciplina"]
    registros = []

    with FileLock(LOCK_PATH, timeout=15):
        wb = load_workbook(EXCEL_PATH, data_only=True)
        ws = wb["Base_Justificativas"]

        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row[0]:
                continue
            if row[2] == projeto_atual and row[3] == sub_atual:
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
    return jsonify({"sucesso": True, "registros": registros})


@app.route("/api/registros", methods=["POST"])
@login_requerido
def adicionar_registro():
    dados = request.get_json() or {}
    projeto_atual = session["projeto"]
    sub_atual = session["subdisciplina"]
    usuario_atual = session["usuario"]

    colaborador = (dados.get("colaborador") or "").strip()
    mes_operacao = (dados.get("mes_operacao") or "").strip()
    semana_operacao = (dados.get("semana_operacao") or "").strip()
    justificativa = (dados.get("justificativa") or "").strip()

    if not all([colaborador, mes_operacao, semana_operacao, justificativa]):
        return jsonify({"sucesso": False, "mensagem": "Preencha os campos obrigatórios."}), 400

    try:
        with FileLock(LOCK_PATH, timeout=20):
            wb = load_workbook(EXCEL_PATH)
            ws = wb["Base_Justificativas"]

            novo_id = datetime.now().strftime("%Y%m%d%H%M%S%f")[:18]
            data_hora_fmt = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

            nova_linha = [
                novo_id,
                data_hora_fmt,
                projeto_atual,
                sub_atual,
                usuario_atual,
                colaborador,
                mes_operacao,
                semana_operacao,
                justificativa
            ]
            ws.append(nova_linha)
            wb.save(EXCEL_PATH)
            wb.close()

        return jsonify({"sucesso": True})
    except PermissionError:
        return jsonify({
            "sucesso": False,
            "mensagem": "O arquivo Excel está aberto no servidor. Feche-o e tente novamente."
        }), 423
    except Exception as e:
        return jsonify({"sucesso": False, "mensagem": f"Erro interno: {str(e)}"}), 500


@app.route("/api/registros/<registro_id>", methods=["DELETE"])
@login_requerido
def deletar_registro(registro_id):
    projeto_atual = session["projeto"]
    sub_atual = session["subdisciplina"]
    try:
        with FileLock(LOCK_PATH, timeout=20):
            wb = load_workbook(EXCEL_PATH)
            ws = wb["Base_Justificativas"]

            linha_para_deletar = None
            for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                if str(row[0]) == str(registro_id) and str(row[2]) == projeto_atual and str(row[3]) == sub_atual:
                    linha_para_deletar = idx
                    break

            if linha_para_deletar:
                ws.delete_rows(linha_para_deletar, 1)
                wb.save(EXCEL_PATH)
                wb.close()
                return jsonify({"sucesso": True})

            wb.close()
            return jsonify({"sucesso": False, "mensagem": "Registro não encontrado."}), 404
    except PermissionError:
        return jsonify({
            "sucesso": False,
            "mensagem": "Feche o arquivo Excel aberto no servidor antes de excluir."
        }), 423


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)
