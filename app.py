from flask import Flask, render_template, request, send_file, jsonify, redirect
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
import io
import base64
import fitz
import unicodedata
import re
import sqlite3
import csv
import html
from datetime import datetime

app = Flask(__name__)
DB = "validacao_crachas.db"

REFEICOES_PADRAO = [
    "CAFÉ DA MANHÃ",
    "ALMOÇO",
    "LANCHE",
    "JANTA"
]


# =========================
# BANCO DE DADOS
# =========================

def init_db():
    conn = sqlite3.connect(DB)
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS atletas (
            codigo_qr TEXT PRIMARY KEY,
            nome TEXT,
            tipo_pessoa TEXT,
            categoria TEXT,
            pagina INTEGER,
            escola TEXT,
            status TEXT DEFAULT 'PENDENTE',
            checkin_hora TEXT
        )
    """)

    cur.execute("PRAGMA table_info(atletas)")
    colunas_atletas = [col[1] for col in cur.fetchall()]
    if "tipo_pessoa" not in colunas_atletas:
        cur.execute("ALTER TABLE atletas ADD COLUMN tipo_pessoa TEXT")

    cur.execute("""
        CREATE TABLE IF NOT EXISTS checkins (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            codigo_qr TEXT,
            tipo TEXT,
            horario TEXT,
            UNIQUE(codigo_qr, tipo)
        )
    """)

    conn.commit()
    conn.close()


def salvar_atleta(codigo_qr, nome, tipo_pessoa, categoria, pagina, escola=""):
    codigo_qr = str(codigo_qr).strip()

    if not codigo_qr or codigo_qr.lower() == "nan":
        return

    conn = sqlite3.connect(DB)
    cur = conn.cursor()

    cur.execute("""
        INSERT OR REPLACE INTO atletas
        (codigo_qr, nome, tipo_pessoa, categoria, pagina, escola, status, checkin_hora)
        VALUES (
            ?, ?, ?, ?, ?, ?,
            COALESCE((SELECT status FROM atletas WHERE codigo_qr = ?), 'PENDENTE'),
            COALESCE((SELECT checkin_hora FROM atletas WHERE codigo_qr = ?), NULL)
        )
    """, (codigo_qr, nome, tipo_pessoa, categoria, pagina, escola, codigo_qr, codigo_qr))

    conn.commit()
    conn.close()


def buscar_atleta_por_codigo(codigo_qr):
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("SELECT * FROM atletas WHERE codigo_qr = ?", (str(codigo_qr).strip(),))
    atleta = cur.fetchone()

    conn.close()
    return atleta


def registrar_entrada_geral(codigo_qr):
    atleta = buscar_atleta_por_codigo(codigo_qr)

    if not atleta:
        return False, "ATLETA NÃO ENCONTRADO"

    if atleta["status"] == "ENTROU":
        return False, "JÁ ENTROU"

    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    conn = sqlite3.connect(DB)
    cur = conn.cursor()

    cur.execute("""
        UPDATE atletas
        SET status = 'ENTROU', checkin_hora = ?
        WHERE codigo_qr = ?
    """, (agora, codigo_qr))

    conn.commit()
    conn.close()

    return True, "ENTRADA REGISTRADA"


def registrar_checkin_tipo(codigo_qr, tipo):
    codigo_qr = str(codigo_qr).strip()
    tipo = str(tipo).strip().upper()

    atleta = buscar_atleta_por_codigo(codigo_qr)

    if not atleta:
        return {
            "ok": False,
            "status": "NAO_ENCONTRADO",
            "mensagem": "ATLETA NÃO ENCONTRADO",
            "codigo_qr": codigo_qr
        }

    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    conn = sqlite3.connect(DB)
    cur = conn.cursor()

    try:
        cur.execute("""
            INSERT INTO checkins (codigo_qr, tipo, horario)
            VALUES (?, ?, ?)
        """, (codigo_qr, tipo, agora))

        conn.commit()
        conn.close()

        return {
            "ok": True,
            "status": "REGISTRADO",
            "mensagem": f"{tipo} REGISTRADO",
            "codigo_qr": codigo_qr,
            "nome": atleta["nome"],
            "tipo_pessoa": atleta["tipo_pessoa"],
            "categoria": atleta["categoria"],
            "escola": atleta["escola"],
            "horario": agora
        }

    except sqlite3.IntegrityError:
        cur.execute("""
            SELECT horario FROM checkins
            WHERE codigo_qr = ? AND tipo = ?
        """, (codigo_qr, tipo))
        row = cur.fetchone()

        conn.close()

        return {
            "ok": False,
            "status": "DUPLICADO",
            "mensagem": f"JÁ REGISTRADO EM {tipo}",
            "codigo_qr": codigo_qr,
            "nome": atleta["nome"],
            "tipo_pessoa": atleta["tipo_pessoa"],
            "categoria": atleta["categoria"],
            "escola": atleta["escola"],
            "horario": row[0] if row else ""
        }


# =========================
# FUNÇÕES DE TEXTO
# =========================

def limpar_texto(txt):
    txt = str(txt).upper().strip()
    txt = unicodedata.normalize("NFD", txt)
    txt = "".join(c for c in txt if unicodedata.category(c) != "Mn")
    txt = re.sub(r"[^A-Z0-9 ]", " ", txt)
    txt = re.sub(r"\s+", " ", txt)
    return txt.strip()


def limpar_codigo(valor):
    codigo = str(valor).strip()
    codigo = re.sub(r"\D", "", codigo)
    return codigo


def normalizar_sexo(valor):
    sexo = limpar_texto(valor)

    if sexo in ["M", "MASC", "MASCULINO"]:
        return "MASCULINO"

    if sexo in ["F", "FEM", "FEMININO"]:
        return "FEMININO"

    return None


def calcular_categoria(data_nascimento):
    nascimento = pd.to_datetime(data_nascimento, dayfirst=True, errors="coerce")

    if pd.isna(nascimento):
        return None

    ano = nascimento.year

    if 2009 <= ano <= 2011:
        return "15 A 17 ANOS"

    if 2012 <= ano <= 2014:
        return "12 A 14 ANOS"

    return None


def encontrar_coluna(df, nomes_possiveis):
    for col in df.columns:
        col_limpa = limpar_texto(col)
        for nome in nomes_possiveis:
            if limpar_texto(nome) in col_limpa:
                return col
    return None


def buscar_linha_por_nome(df, texto_pagina, col_nome):
    texto_pagina_limpo = limpar_texto(texto_pagina)

    melhor_idx = None
    melhor_row = None
    melhor_pontuacao = 0

    for idx, row in df.iterrows():
        nome_excel = limpar_texto(row[col_nome])

        if not nome_excel:
            continue

        partes = nome_excel.split()

        if len(partes) < 2:
            continue

        acertos = sum(1 for parte in partes if parte in texto_pagina_limpo)
        pontuacao = acertos / len(partes)

        primeiro_nome_ok = partes[0] in texto_pagina_limpo
        segundo_nome_ok = partes[1] in texto_pagina_limpo

        if primeiro_nome_ok and segundo_nome_ok and pontuacao > melhor_pontuacao:
            melhor_pontuacao = pontuacao
            melhor_idx = idx
            melhor_row = row

    if melhor_pontuacao >= 0.50:
        return melhor_idx, melhor_row

    return None, None


def extrair_codigo_qr_do_texto(texto_pagina):
    texto_original = str(texto_pagina)
    candidatos = re.findall(r"\b\d{6,20}\b", texto_original)

    if candidatos:
        return candidatos[0]

    texto_limpo = limpar_texto(texto_pagina)
    candidatos = re.findall(r"\b\d{6,20}\b", texto_limpo)

    if candidatos:
        return candidatos[0]

    return None




def extrair_nome_do_pdf(texto_pagina):
    linhas = [str(l).strip() for l in str(texto_pagina).splitlines() if str(l).strip()]

    ignorar = [
        "JOGOS", "ESCOLARES", "RORAIMA", "WWW", "ATLETA",
        "TÉCNICO", "TECNICO", "OFICIAL", "CHEFE", "DELEGAÇÃO",
        "DELEGACAO", "FUTSAL", "VOLEI", "VÔLEI", "ATLETISMO",
        "CICLISMO", "BADMINTON", "CAFÉ", "CAFE", "ALMOÇO",
        "ALMOCO", "LANCHE", "JANTA"
    ]

    for linha in linhas:
        limpa = limpar_texto(linha)

        if not limpa:
            continue

        if any(palavra in limpa for palavra in ignorar):
            continue

        linha_sem_numero = re.sub(r"\b\d{6,20}\b", "", linha).strip()

        if len(linha_sem_numero.split()) >= 2:
            return linha_sem_numero.upper()

    return "NOME NÃO IDENTIFICADO"


def extrair_numero_credencial(texto_pagina):
    texto_original = str(texto_pagina)

    candidatos = re.findall(r"\b\d{6,20}\b", texto_original)

    if candidatos:
        return candidatos[0]

    texto_limpo = limpar_texto(texto_pagina)
    candidatos = re.findall(r"\b\d{6,20}\b", texto_limpo)

    if candidatos:
        return candidatos[0]

    return ""


def montar_categoria(row, col_nome, col_sexo, col_data, linha_excel):
    nome = str(row[col_nome]).strip()

    sexo = normalizar_sexo(row[col_sexo])
    categoria = calcular_categoria(row[col_data])

    if not sexo:
        raise Exception(f"Linha {linha_excel}: sexo inválido para {nome}")

    if not categoria:
        raise Exception(f"Linha {linha_excel}: data de nascimento inválida ou fora da categoria para {nome}")

    return f"{categoria} {sexo}"



def montar_credencial(row, col_credencial, linha_excel):
    if not col_credencial:
        raise Exception("Não encontrei a coluna da CREDENCIAL na planilha. Use uma coluna chamada CREDENCIAL, Nº CREDENCIAL, NUMERO CREDENCIAL ou CODIGO CREDENCIAL.")

    valor = valor_linha(row, col_credencial, "")
    if not valor_valido(valor):
        raise Exception(f"Linha {linha_excel}: número da credencial vazio ou inválido.")

    if re.fullmatch(r"\d+\.0", str(valor)):
        valor = str(valor).replace(".0", "")

    return str(valor).strip().upper()


def valor_linha(row, coluna, padrao=""):
    if not coluna:
        return padrao
    valor = row.get(coluna, padrao)
    if pd.isna(valor):
        return padrao
    valor = str(valor).strip()
    if valor.lower() == "nan":
        return padrao
    return valor


def valor_valido(valor):
    texto = limpar_texto(valor)
    return bool(texto and texto not in ["---", "-", "NA", "NAN", "NONE", "NULL"])


def identificar_tipo_pessoa(row, col_funcao=None, col_tipo_usuario=None, texto_pagina=""):
    funcao = valor_linha(row, col_funcao, "")
    tipo_usuario = valor_linha(row, col_tipo_usuario, "")
    texto_pdf = limpar_texto(texto_pagina)

    if valor_valido(funcao):
        return limpar_texto(funcao)

    if "CHEFE DE DELEGACAO" in texto_pdf:
        return "CHEFE DE DELEGAÇÃO"
    if "TECNICO" in texto_pdf:
        return "TÉCNICO"
    if "OFICIAL" in texto_pdf:
        return "OFICIAL"

    if "PRESTADOR" in limpar_texto(tipo_usuario):
        return "DIRIGENTE"

    return "ATLETA"


def montar_texto_cracha(row, col_nome, col_sexo, col_data, col_escola, col_funcao, col_tipo_usuario, linha_excel, texto_pagina=""):
    tipo_pessoa = identificar_tipo_pessoa(row, col_funcao, col_tipo_usuario, texto_pagina)
    escola = valor_linha(row, col_escola, "")

    if tipo_pessoa != "ATLETA":
        if not valor_valido(escola):
            nome = str(row[col_nome]).strip()
            raise Exception(f"Linha {linha_excel}: escola inválida para {nome}")
        return escola.upper(), tipo_pessoa, escola

    if not col_sexo:
        raise Exception("Não encontrei a coluna SEXO na planilha para calcular categoria dos atletas.")
    if not col_data:
        raise Exception("Não encontrei a coluna DATA NASCIMENTO na planilha para calcular categoria dos atletas.")

    categoria = montar_categoria(row, col_nome, col_sexo, col_data, linha_excel)
    return categoria, tipo_pessoa, escola


def texto_pagina_fitz(pdf_bytes, pagina_index):
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        texto = doc[pagina_index].get_text()
        doc.close()
        return texto or ""
    except Exception:
        return ""


# =========================
# PDF / BASE
# =========================

def criar_pdf(excel_file, pdf_file, pos_x, pos_y, rotacao, fonte, somente_primeira_pagina=False):
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip().str.upper()

    col_nome = encontrar_coluna(df, ["NOME"])
    col_sexo = encontrar_coluna(df, ["SEXO"])
    col_data = encontrar_coluna(df, ["DATA NASCIMENTO", "DATA DE NASCIMENTO", "NASCIMENTO", "DATA"])
    col_escola = encontrar_coluna(df, ["ESCOLA"])
    col_cpf = encontrar_coluna(df, ["CPF", "CODIGO", "CÓDIGO", "ID"])
    col_funcao = encontrar_coluna(df, ["FUNCAO", "FUNÇÃO", "CARGO"])
    col_tipo_usuario = encontrar_coluna(df, ["TIPO USUARIO", "TIPO USUÁRIO", "TIPO"])
    col_funcao = encontrar_coluna(df, ["FUNCAO", "FUNÇÃO", "CARGO"])
    col_tipo_usuario = encontrar_coluna(df, ["TIPO USUARIO", "TIPO USUÁRIO", "TIPO"])

    if not col_nome:
        raise Exception("Não encontrei a coluna NOME na planilha.")


    pdf_bytes = pdf_file.read()
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()

    total_paginas = 1 if somente_primeira_pagina else len(reader.pages)
    erros = []

    for i in range(total_paginas):
        page = reader.pages[i]

        texto_pypdf = page.extract_text() or ""
        texto_fitz = texto_pagina_fitz(pdf_bytes, i)
        texto_pagina = texto_pypdf + "\n" + texto_fitz

        codigo_qr = extrair_codigo_qr_do_texto(texto_pagina)

        idx_excel, row = buscar_linha_por_nome(df, texto_pagina, col_nome)

        if row is None:
            erros.append(f"Página {i + 1}: não encontrei o atleta do crachá na planilha.")
            continue

        if not codigo_qr and col_cpf:
            codigo_qr = limpar_codigo(row[col_cpf])

        linha_excel = idx_excel + 2

        try:
            texto_cracha, tipo_pessoa, escola = montar_texto_cracha(
                row,
                col_nome,
                col_sexo,
                col_data,
                col_escola,
                col_funcao,
                col_tipo_usuario,
                linha_excel,
                texto_pagina
            )

        except Exception as e:
            erros.append(str(e))
            continue

        nome = str(row[col_nome]).strip()

        salvar_atleta(codigo_qr, nome, tipo_pessoa, texto_cracha, i + 1, escola)

        packet = io.BytesIO()
        largura = float(page.mediabox.width)
        altura = float(page.mediabox.height)

        can = canvas.Canvas(packet, pagesize=(largura, altura))
        can.setFont("Helvetica-Bold", fonte)
        can.setFillColorRGB(1, 1, 1)  # texto branco no crachá

        can.saveState()
        can.translate(pos_x, pos_y)
        can.rotate(rotacao)
        can.drawCentredString(0, 0, texto_cracha)
        can.restoreState()

        can.save()
        packet.seek(0)

        overlay = PdfReader(packet)

        if len(overlay.pages) > 0:
            page.merge_page(overlay.pages[0])

        writer.add_page(page)

    if erros:
        raise Exception("\n".join(erros[:30]))

    output = io.BytesIO()
    writer.write(output)
    output.seek(0)

    return output


def montar_base_validacao(excel_file, pdf_file):
    df = pd.read_excel(excel_file)
    df.columns = df.columns.str.strip().str.upper()

    col_nome = encontrar_coluna(df, ["NOME"])
    col_sexo = encontrar_coluna(df, ["SEXO"])
    col_data = encontrar_coluna(df, ["DATA NASCIMENTO", "DATA DE NASCIMENTO", "NASCIMENTO", "DATA"])
    col_escola = encontrar_coluna(df, ["ESCOLA"])
    col_cpf = encontrar_coluna(df, ["CPF", "CODIGO", "CÓDIGO", "ID"])
    col_funcao = encontrar_coluna(df, ["FUNCAO", "FUNÇÃO", "CARGO"])
    col_tipo_usuario = encontrar_coluna(df, ["TIPO USUARIO", "TIPO USUÁRIO", "TIPO"])

    pdf_bytes = pdf_file.read()
    reader = PdfReader(io.BytesIO(pdf_bytes))

    base = []
    erros = []

    for i, page in enumerate(reader.pages):
        texto_pypdf = page.extract_text() or ""
        texto_fitz = texto_pagina_fitz(pdf_bytes, i)
        texto_pagina = texto_pypdf + "\n" + texto_fitz

        codigo_qr = extrair_codigo_qr_do_texto(texto_pagina)
        idx_excel, row = buscar_linha_por_nome(df, texto_pagina, col_nome)

        if row is None:
            erros.append({
                "pagina": i + 1,
                "codigo_qr": codigo_qr,
                "erro": "Atleta não encontrado na planilha"
            })
            continue

        if not codigo_qr and col_cpf:
            codigo_qr = limpar_codigo(row[col_cpf])

        linha_excel = idx_excel + 2

        try:
            texto_cracha, tipo_pessoa, escola = montar_texto_cracha(
                row,
                col_nome,
                col_sexo,
                col_data,
                col_escola,
                col_funcao,
                col_tipo_usuario,
                linha_excel,
                texto_pagina
            )
        except Exception as e:
            erros.append({
                "pagina": i + 1,
                "codigo_qr": codigo_qr,
                "erro": str(e)
            })
            continue

        nome = str(row[col_nome]).strip()

        salvar_atleta(codigo_qr, nome, tipo_pessoa, texto_cracha, i + 1, escola)

        base.append({
            "pagina": i + 1,
            "codigo_qr": codigo_qr,
            "nome": nome,
            "tipo_pessoa": tipo_pessoa,
            "categoria": texto_cracha,
            "escola": escola
        })

    return base, erros



# =========================
# LAYOUT BONITO / UI
# =========================

def h(valor):
    return html.escape(str(valor or ""))


def layout(titulo, conteudo, ativo=""):
    itens = [
        ("/dashboard", "🏠", "Início", "dashboard"),
        ("/", "🪪", "Gerar Crachás", "gerar"),
        ("/scanner", "📷", "Scanner", "scanner"),
        ("/painel", "✅", "Painel", "painel"),
        ("/relatorio", "📊", "Relatórios", "relatorio"),
    ]

    nav = ""
    for url, icone, nome, chave in itens:
        classe = "active" if ativo == chave else ""
        nav += f'<a class="nav-item {classe}" href="{url}"><span>{icone}</span>{nome}</a>'

    return f"""
    <!doctype html>
    <html lang="pt-br">
    <head>
        <meta charset="utf-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{h(titulo)}</title>
        <style>
            :root {{
                --bg:#f4f7fb;
                --card:#ffffff;
                --dark:#0f172a;
                --muted:#64748b;
                --line:#e5e7eb;
                --blue:#2563eb;
                --green:#16a34a;
                --red:#dc2626;
                --amber:#f59e0b;
                --shadow:0 10px 25px rgba(15,23,42,.08);
                --radius:18px;
            }}
            * {{ box-sizing:border-box; }}
            body {{
                margin:0;
                font-family:Inter, Arial, sans-serif;
                background:var(--bg);
                color:var(--dark);
            }}
            .app {{ display:flex; min-height:100vh; }}
            .sidebar {{
                width:260px;
                background:linear-gradient(180deg,#0f172a,#111827);
                color:white;
                padding:22px 16px;
                position:fixed;
                left:0; top:0; bottom:0;
            }}
            .brand {{
                display:flex; align-items:center; gap:12px;
                padding:10px 10px 24px 10px;
                border-bottom:1px solid rgba(255,255,255,.12);
                margin-bottom:18px;
            }}
            .logo {{
                width:44px; height:44px; border-radius:14px;
                display:grid; place-items:center;
                background:linear-gradient(135deg,#22c55e,#06b6d4);
                font-size:23px;
                box-shadow:0 10px 25px rgba(34,197,94,.25);
            }}
            .brand strong {{ display:block; font-size:18px; }}
            .brand small {{ color:#cbd5e1; }}
            .nav-item {{
                display:flex; align-items:center; gap:12px;
                color:#cbd5e1;
                text-decoration:none;
                padding:13px 14px;
                border-radius:14px;
                margin-bottom:8px;
                font-weight:700;
            }}
            .nav-item:hover, .nav-item.active {{
                color:white;
                background:rgba(255,255,255,.12);
            }}
            .content {{
                margin-left:260px;
                width:calc(100% - 260px);
                padding:26px;
            }}
            .topbar {{
                display:flex; align-items:center; justify-content:space-between;
                margin-bottom:22px;
            }}
            .topbar h1 {{ margin:0; font-size:30px; letter-spacing:-.03em; }}
            .topbar p {{ margin:6px 0 0; color:var(--muted); }}
            .actions {{ display:flex; gap:10px; flex-wrap:wrap; }}
            .btn {{
                display:inline-flex; align-items:center; justify-content:center; gap:8px;
                padding:12px 16px;
                border-radius:12px;
                border:0;
                text-decoration:none;
                font-weight:800;
                cursor:pointer;
                background:#e2e8f0;
                color:#0f172a;
            }}
            .btn.primary {{ background:var(--green); color:white; }}
            .btn.blue {{ background:var(--blue); color:white; }}
            .btn.dark {{ background:var(--dark); color:white; }}
            .btn.red {{ background:var(--red); color:white; }}
            .grid {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(210px,1fr)); gap:16px; }}
            .card {{
                background:var(--card);
                border:1px solid var(--line);
                border-radius:var(--radius);
                padding:20px;
                box-shadow:var(--shadow);
            }}
            .metric-label {{ color:var(--muted); font-weight:800; font-size:13px; text-transform:uppercase; letter-spacing:.04em; }}
            .metric-value {{ font-size:38px; font-weight:900; margin-top:8px; letter-spacing:-.04em; }}
            .panel {{
                background:white;
                border:1px solid var(--line);
                border-radius:var(--radius);
                box-shadow:var(--shadow);
                padding:20px;
                margin-top:18px;
            }}
            table {{ border-collapse:collapse; width:100%; background:white; overflow:hidden; border-radius:14px; }}
            th,td {{ padding:12px; border-bottom:1px solid var(--line); text-align:left; font-size:14px; }}
            th {{ background:#0f172a; color:white; position:sticky; top:0; z-index:1; }}
            tr:hover td {{ background:#f8fafc; }}
            input, select {{
                width:100%; padding:12px; border:1px solid #cbd5e1; border-radius:12px;
                font-size:15px; background:white;
            }}
            label {{ font-weight:800; font-size:13px; color:#334155; }}
            .filters {{ display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:12px; align-items:end; }}
            .status-ok {{ color:var(--green); font-weight:900; }}
            .status-pendente {{ color:var(--red); font-weight:900; }}
            .pill {{ display:inline-block; padding:6px 10px; border-radius:999px; background:#e2e8f0; font-weight:800; font-size:12px; }}
            .pill.green {{ background:#dcfce7; color:#166534; }}
            .pill.amber {{ background:#fef3c7; color:#92400e; }}
            .pill.red {{ background:#fee2e2; color:#991b1b; }}
            .table-wrap {{ max-height:70vh; overflow:auto; border-radius:14px; border:1px solid var(--line); }}
            @media (max-width: 850px) {{
                .app {{ display:block; }}
                .sidebar {{ position:relative; width:100%; bottom:auto; }}
                .content {{ margin-left:0; width:100%; padding:16px; }}
                .topbar {{ display:block; }}
                .actions {{ margin-top:12px; }}
                .nav-item {{ display:inline-flex; margin-right:6px; }}
            }}
        </style>
    </head>
    <body>
        <div class="app">
            <aside class="sidebar">
                <div class="brand">
                    <div class="logo">🏆</div>
                    <div>
                        <strong>Sistema de Crachá</strong>
                        <small>Jogos Escolares</small>
                    </div>
                </div>
                {nav}
            </aside>
            <main class="content">
                {conteudo}
            </main>
        </div>
    </body>
    </html>
    """


def get_metricas():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) total FROM atletas")
    total = cur.fetchone()["total"]

    cur.execute("SELECT COUNT(*) total FROM atletas WHERE status = 'ENTROU'")
    entrada = cur.fetchone()["total"]

    cur.execute("SELECT COUNT(*) total FROM checkins")
    refeicoes = cur.fetchone()["total"]

    cur.execute("SELECT tipo, COUNT(*) total FROM checkins GROUP BY tipo")
    por_tipo = {r["tipo"]: r["total"] for r in cur.fetchall()}

    cur.execute("""
        SELECT c.codigo_qr, c.tipo, c.horario, a.nome, a.escola, a.categoria
        FROM checkins c
        LEFT JOIN atletas a ON a.codigo_qr = c.codigo_qr
        ORDER BY c.id DESC
        LIMIT 8
    """)
    ultimos = cur.fetchall()

    conn.close()
    return total, entrada, refeicoes, por_tipo, ultimos


# =========================
# ROTAS PRINCIPAIS
# =========================

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/dashboard")
def dashboard():
    total, entrada, refeicoes, por_tipo, ultimos = get_metricas()
    cards_refeicoes = ""
    for r in REFEICOES_PADRAO:
        cards_refeicoes += f"""
        <div class="card">
            <div class="metric-label">{h(r)}</div>
            <div class="metric-value">{por_tipo.get(r, 0)}</div>
        </div>
        """

    linhas = ""
    for u in ultimos:
        linhas += f"""
        <tr>
            <td>{h(u['horario'])}</td>
            <td><span class="pill green">{h(u['tipo'])}</span></td>
            <td>{h(u['codigo_qr'])}</td>
            <td>{h(u['nome'])}</td>
            <td>{h(u['escola'])}</td>
        </tr>
        """

    conteudo = f"""
    <div class="topbar">
        <div>
            <h1>Dashboard</h1>
            <p>Visão rápida do evento, refeições e últimos registros.</p>
        </div>
        <div class="actions">
            <a class="btn primary" href="/scanner">📷 Abrir Scanner</a>
            <a class="btn blue" href="/relatorio">📊 Relatório</a>
            <a class="btn dark" href="/painel">✅ Painel</a>
        </div>
    </div>

    <div class="grid">
        <div class="card"><div class="metric-label">Total de Atletas</div><div class="metric-value">{total}</div></div>
        <div class="card"><div class="metric-label">Entrada Geral</div><div class="metric-value">{entrada}</div></div>
        <div class="card"><div class="metric-label">Registros de Refeições</div><div class="metric-value">{refeicoes}</div></div>
        {cards_refeicoes}
    </div>

    <div class="panel">
        <h2>Últimos registros</h2>
        <div class="table-wrap">
        <table>
            <tr><th>Horário</th><th>Tipo</th><th>QR</th><th>Nome</th><th>Escola</th></tr>
            {linhas if linhas else '<tr><td colspan="5">Nenhum registro ainda.</td></tr>'}
        </table>
        </div>
    </div>
    """
    return layout("Dashboard", conteudo, "dashboard")


@app.route("/preview", methods=["POST"])
def preview():
    try:
        excel = request.files["excel"]
        pdf = request.files["pdf"]

        pos_x = int(request.form.get("pos_x", 118))
        pos_y = int(request.form.get("pos_y", 300))
        rotacao = int(request.form.get("rotacao", 90))
        fonte = int(request.form.get("fonte", 10))
        pdf_preview = criar_pdf(excel, pdf, pos_x, pos_y, rotacao, fonte, somente_primeira_pagina=True)

        doc = fitz.open(stream=pdf_preview.getvalue(), filetype="pdf")
        page = doc[0]
        pix = page.get_pixmap(matrix=fitz.Matrix(1.2, 1.2), alpha=False)
        img_base64 = base64.b64encode(pix.tobytes("png")).decode("utf-8")
        doc.close()

        return jsonify({"ok": True, "imagem": f"data:image/png;base64,{img_base64}"})
    except Exception as e:
        return jsonify({"ok": False, "erro": str(e)})


@app.route("/gerar", methods=["POST"])
def gerar():
    try:
        excel = request.files["excel"]
        pdf = request.files["pdf"]

        pos_x = int(request.form.get("pos_x", 118))
        pos_y = int(request.form.get("pos_y", 300))
        rotacao = int(request.form.get("rotacao", 90))
        fonte = int(request.form.get("fonte", 10))
        output = criar_pdf(excel, pdf, pos_x, pos_y, rotacao, fonte, somente_primeira_pagina=False)

        return send_file(output, as_attachment=True, download_name="crachas_final.pdf", mimetype="application/pdf")
    except Exception as e:
        return f"Erro ao gerar PDF:<br><pre>{h(e)}</pre>"


@app.route("/gerar-excel-credenciais", methods=["POST"])
def gerar_excel_credenciais():
    try:
        excel = request.files["excel"]
        pdf = request.files["pdf"]

        df = pd.read_excel(excel)
        df.columns = df.columns.str.strip().str.upper()

        col_nome = encontrar_coluna(df, ["NOME"])
        col_cpf = encontrar_coluna(df, ["CPF"])

        if not col_nome:
            raise Exception("Não encontrei a coluna NOME na planilha.")

        pdf_bytes = pdf.read()
        reader = PdfReader(io.BytesIO(pdf_bytes))

        registros = []

        for i, page in enumerate(reader.pages):
            texto_pypdf = page.extract_text() or ""
            texto_fitz = texto_pagina_fitz(pdf_bytes, i)
            texto_pagina = texto_pypdf + "\n" + texto_fitz

            numero_credencial = extrair_numero_credencial(texto_pagina)
            nome_lido_pdf = extrair_nome_do_pdf(texto_pagina)

            idx_excel, row = buscar_linha_por_nome(df, texto_pagina, col_nome)

            if row is not None:
                cpf = ""
                if col_cpf:
                    cpf = valor_linha(row, col_cpf, "")

                novo = {
    "PAGINA_PDF": i + 1,
    "NUMERO_CREDENCIAL": numero_credencial,
    "CPF": cpf
}

                dados = row.to_dict()
                for coluna, valor in dados.items():
                    if coluna not in novo:
                        novo[coluna] = valor

                registros.append(novo)

            else:
                registros.append({
    "PAGINA_PDF": i + 1,
    "NUMERO_CREDENCIAL": numero_credencial,
    "CPF": ""
})

        df_saida = pd.DataFrame(registros)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_saida.to_excel(writer, index=False, sheet_name="Credenciais")

        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="credenciais_extraidas.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return f"Erro ao gerar Excel de credenciais:<br><pre>{h(e)}</pre>"


@app.route("/validar-base", methods=["POST"])
def validar_base():
    try:
        excel = request.files["excel"]
        pdf = request.files["pdf"]
        base, erros = montar_base_validacao(excel, pdf)

        linhas = ""
        for item in base:
            codigo = h(item["codigo_qr"])
            linhas += f"""
            <tr>
                <td>{h(item['pagina'])}</td>
                <td>{codigo}</td>
                <td>{h(item['nome'])}</td>
                <td>{h(item.get('tipo_pessoa', ''))}</td>
                <td>{h(item['categoria'])}</td>
                <td>{h(item['escola'])}</td>
                <td><a class="btn blue" href="/atleta/{codigo}" target="_blank">Abrir</a></td>
            </tr>
            """

        linhas_erros = ""
        for erro in erros:
            linhas_erros += f"""
            <tr>
                <td>{h(erro['pagina'])}</td>
                <td>{h(erro['codigo_qr'])}</td>
                <td>{h(erro['erro'])}</td>
            </tr>
            """

        erros_html = ""
        if erros:
            erros_html = f"""
            <div class="panel">
                <h2>Erros encontrados: {len(erros)}</h2>
                <div class="table-wrap"><table><tr><th>Página</th><th>QR</th><th>Erro</th></tr>{linhas_erros}</table></div>
            </div>
            """

        conteudo = f"""
        <div class="topbar">
            <div><h1>Base de Validação</h1><p>Atletas cadastrados no banco para uso no scanner.</p></div>
            <div class="actions"><a class="btn primary" href="/scanner">📷 Scanner</a><a class="btn blue" href="/relatorio">Relatório</a></div>
        </div>
        <div class="grid"><div class="card"><div class="metric-label">Atletas encontrados</div><div class="metric-value">{len(base)}</div></div></div>
        <div class="panel">
            <h2>Atletas</h2>
            <div class="table-wrap"><table><tr><th>Página</th><th>QR</th><th>Nome</th><th>Tipo</th><th>Texto no Crachá</th><th>Escola</th><th>Link</th></tr>{linhas}</table></div>
        </div>
        {erros_html}
        """
        return layout("Base de Validação", conteudo, "painel")
    except Exception as e:
        return layout("Erro", f"<div class='panel'><h1>Erro ao validar base</h1><pre>{h(e)}</pre></div>", "painel")


# =========================
# ATLETA / ENTRADA GERAL
# =========================

@app.route("/atleta/<codigo_qr>")
def atleta(codigo_qr):
    atleta = buscar_atleta_por_codigo(codigo_qr)

    if not atleta:
        conteudo = f"""
        <div class="topbar"><div><h1>Atleta não encontrado</h1><p>Código lido: {h(codigo_qr)}</p></div></div>
        <div class="panel" style="border-left:8px solid var(--red);">
            <h2 style="color:var(--red);">ATLETA NÃO ENCONTRADO</h2>
            <p>Verifique se a base foi gerada ou se o QR foi lido corretamente.</p>
            <a class="btn primary" href="/scanner">Voltar ao scanner</a>
        </div>
        """
        return layout("Atleta não encontrado", conteudo, "scanner")

    status = atleta["status"]
    pill = '<span class="pill green">LIBERADO PARA ENTRAR</span>' if status == "PENDENTE" else '<span class="pill amber">JÁ ENTROU</span>'

    if status == "PENDENTE":
        botao = f"""
        <form action="/checkin/{h(codigo_qr)}" method="POST">
            <button class="btn primary" type="submit">CONFIRMAR ENTRADA GERAL</button>
        </form>
        """
    else:
        botao = f"<p>Entrada registrada em: <strong>{h(atleta['checkin_hora'])}</strong></p>"

    conteudo = f"""
    <div class="topbar">
        <div><h1>{h(atleta['nome'])}</h1><p>Dados do crachá e entrada geral.</p></div>
        <div class="actions"><a class="btn primary" href="/scanner">📷 Scanner</a><a class="btn blue" href="/relatorio?busca={h(codigo_qr)}">Ver histórico</a></div>
    </div>
    <div class="panel">
        <div class="grid">
            <div class="card"><div class="metric-label">Código QR</div><div style="font-size:26px;font-weight:900;">{h(atleta['codigo_qr'])}</div></div>
            <div class="card"><div class="metric-label">Tipo</div><div style="font-size:22px;font-weight:900;">{h(atleta['tipo_pessoa'])}</div></div>
            <div class="card"><div class="metric-label">Texto no Crachá</div><div style="font-size:22px;font-weight:900;">{h(atleta['categoria'])}</div></div>
            <div class="card"><div class="metric-label">Escola</div><div style="font-size:18px;font-weight:900;">{h(atleta['escola'])}</div></div>
            <div class="card"><div class="metric-label">Status</div><div style="margin-top:12px;">{pill}</div></div>
        </div>
        <div style="margin-top:20px;">{botao}</div>
    </div>
    """
    return layout("Atleta", conteudo, "painel")


@app.route("/checkin/<codigo_qr>", methods=["POST"])
def checkin(codigo_qr):
    registrar_entrada_geral(codigo_qr)
    return redirect(f"/atleta/{codigo_qr}")


# =========================
# API DO SCANNER
# =========================

@app.route("/api/checkin-refeicao", methods=["POST"])
def api_checkin_refeicao():
    data = request.get_json() or {}
    codigo_qr = data.get("codigo_qr", "")
    tipo = data.get("tipo", "")

    if not codigo_qr or not tipo:
        return jsonify({"ok": False, "mensagem": "Código QR ou tipo vazio."})

    resultado = registrar_checkin_tipo(codigo_qr, tipo)
    return jsonify(resultado)


# =========================
# SCANNER CELULAR
# =========================

@app.route("/scanner")
def scanner():
    opcoes = "".join([f'<option value="{h(r)}">{h(r)}</option>' for r in REFEICOES_PADRAO])

    conteudo = f"""
    <style>
        .scanner-page {{ max-width:780px; margin:auto; }}
        #reader {{ width:100%; background:white; border-radius:22px; overflow:hidden; box-shadow:var(--shadow); }}
        .resultado {{ margin-top:16px; padding:20px; border-radius:18px; font-size:18px; min-height:105px; font-weight:700; }}
        .resultado.ok {{ background:#16a34a; color:white; }}
        .resultado.erro {{ background:#dc2626; color:white; }}
        .resultado.alerta {{ background:#f59e0b; color:#111827; }}
        .scanner-top {{ position:sticky; top:0; z-index:20; background:var(--bg); padding-bottom:10px; }}
        @media (max-width:850px) {{ .sidebar {{ display:none; }} .content {{ padding:12px; }} }}
    </style>
    <div class="scanner-page">
        <div class="scanner-top">
            <div class="topbar">
                <div><h1>Scanner QR Code</h1><p>Escolha o controle e aponte a câmera para o crachá.</p></div>
                <div class="actions"><a class="btn blue" href="/relatorio">Relatório</a></div>
            </div>
            <div class="panel" style="margin-top:0;">
                <label>Tipo de controle</label>
                <select id="tipo">{opcoes}</select>
            </div>
        </div>

        <div id="reader"></div>
        <div id="resultado" class="resultado alerta">Aponte a câmera para o QR Code do crachá.</div>

        <div class="panel">
            <h3>Digitar código manualmente</h3>
            <input id="manual" placeholder="Ex: 061020880">
            <button class="btn primary" style="width:100%;margin-top:10px;" onclick="registrarManual()">Registrar</button>
        </div>
    </div>

    <script src="https://unpkg.com/html5-qrcode"></script>
    <script>
        let ultimoCodigo = "";
        let travado = false;

        function somenteNumero(texto) {{ return String(texto || "").replace(/\D/g, ""); }}

        function tocar(tipo) {{
            try {{
                if (navigator.vibrate) navigator.vibrate(tipo === "ok" ? 120 : [120,80,120]);
                const ctx = new (window.AudioContext || window.webkitAudioContext)();
                const osc = ctx.createOscillator();
                const gain = ctx.createGain();
                osc.connect(gain); gain.connect(ctx.destination);
                osc.frequency.value = tipo === "ok" ? 880 : 220;
                gain.gain.value = .08;
                osc.start();
                setTimeout(() => {{ osc.stop(); ctx.close(); }}, 140);
            }} catch(e) {{}}
        }}

        async function registrar(codigoOriginal) {{
            const codigo = somenteNumero(codigoOriginal);
            const tipo = document.getElementById("tipo").value;
            if (!codigo) {{ mostrar("erro", "Código inválido."); tocar("erro"); return; }}
            if (travado && codigo === ultimoCodigo) return;
            travado = true; ultimoCodigo = codigo;
            mostrar("alerta", "Registrando código " + codigo + "...");

            try {{
                const resp = await fetch("/api/checkin-refeicao", {{
                    method: "POST",
                    headers: {{ "Content-Type": "application/json" }},
                    body: JSON.stringify({{ codigo_qr: codigo, tipo: tipo }})
                }});
                const data = await resp.json();
                if (data.status === "REGISTRADO") {{
                    tocar("ok");
                    mostrar("ok", "<strong>✅ " + data.mensagem + "</strong><br>" + data.nome + "<br>" + data.tipo_pessoa + "<br>" + data.categoria + "<br>" + data.escola + "<br>Horário: " + data.horario);
                }} else if (data.status === "DUPLICADO") {{
                    tocar("erro");
                    mostrar("alerta", "<strong>⚠️ " + data.mensagem + "</strong><br>" + data.nome + "<br>Registrado em: " + data.horario);
                }} else {{
                    tocar("erro");
                    mostrar("erro", "<strong>❌ " + data.mensagem + "</strong><br>Código: " + codigo);
                }}
                setTimeout(() => {{ travado = false; }}, 2500);
            }} catch (e) {{
                tocar("erro");
                mostrar("erro", "Erro ao registrar. Verifique a conexão.");
                setTimeout(() => {{ travado = false; }}, 2500);
            }}
        }}

        function registrarManual() {{
            const codigo = document.getElementById("manual").value;
            registrar(codigo);
            document.getElementById("manual").value = "";
        }}
        function mostrar(tipo, html) {{
            const div = document.getElementById("resultado");
            div.className = "resultado " + tipo;
            div.innerHTML = html;
        }}

        const html5QrCode = new Html5Qrcode("reader");
        Html5Qrcode.getCameras().then(cameras => {{
            if (cameras && cameras.length) {{
                let cameraId = cameras[cameras.length - 1].id;
                html5QrCode.start(cameraId, {{ fps: 10, qrbox: {{ width: 260, height: 260 }} }}, decodedText => {{ registrar(decodedText); }}, errorMessage => {{}});
            }} else {{ mostrar("erro", "Nenhuma câmera encontrada."); }}
        }}).catch(err => {{ mostrar("erro", "Não foi possível acessar a câmera. Use HTTPS/ngrok e permita a câmera."); }});
    </script>
    """
    return layout("Scanner", conteudo, "scanner")


# =========================
# PAINEL
# =========================

@app.route("/painel")
def painel():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) as total FROM atletas")
    total = cur.fetchone()["total"]
    cur.execute("SELECT COUNT(*) as total FROM atletas WHERE status = 'ENTROU'")
    entrou = cur.fetchone()["total"]
    cur.execute("SELECT COUNT(*) as total FROM atletas WHERE status = 'PENDENTE'")
    pendente = cur.fetchone()["total"]
    cur.execute("SELECT COUNT(*) as total FROM checkins")
    total_refeicoes = cur.fetchone()["total"]
    cur.execute("SELECT * FROM atletas ORDER BY nome")
    atletas = cur.fetchall()
    conn.close()

    linhas = ""
    for a in atletas:
        classe = "status-ok" if a["status"] == "ENTROU" else "status-pendente"
        linhas += f"""
        <tr>
            <td>{h(a['codigo_qr'])}</td>
            <td>{h(a['nome'])}</td>
            <td>{h(a['tipo_pessoa'])}</td>
            <td>{h(a['categoria'])}</td>
            <td>{h(a['escola'])}</td>
            <td class="{classe}">{h(a['status'])}</td>
            <td>{h(a['checkin_hora'])}</td>
            <td><a class="btn blue" href="/atleta/{h(a['codigo_qr'])}" target="_blank">Abrir</a></td>
        </tr>
        """

    conteudo = f"""
    <div class="topbar">
        <div><h1>Painel de Check-in</h1><p>Controle geral dos atletas cadastrados.</p></div>
        <div class="actions"><a class="btn primary" href="/scanner">📷 Abrir Scanner</a><a class="btn blue" href="/relatorio">📊 Relatórios</a></div>
    </div>

    <form class="panel" onsubmit="event.preventDefault(); irAtleta();" style="margin-top:0;">
        <label>Buscar por código QR</label>
        <div style="display:flex;gap:10px;margin-top:8px;">
            <input id="codigo" placeholder="Digite ou escaneie o código QR" autofocus>
            <button class="btn dark" type="submit">Buscar</button>
        </div>
    </form>
    <script>
        function irAtleta() {{
            const codigo = document.getElementById("codigo").value.trim();
            if (codigo) window.location.href = "/atleta/" + codigo;
        }}
    </script>

    <div class="grid">
        <div class="card"><div class="metric-label">Total de Atletas</div><div class="metric-value">{total}</div></div>
        <div class="card"><div class="metric-label">Entrada Geral</div><div class="metric-value">{entrou}</div></div>
        <div class="card"><div class="metric-label">Pendentes Geral</div><div class="metric-value">{pendente}</div></div>
        <div class="card"><div class="metric-label">Registros Refeições</div><div class="metric-value">{total_refeicoes}</div></div>
    </div>

    <div class="panel">
        <h2>Atletas</h2>
        <div class="table-wrap">
        <table>
            <tr><th>Código QR</th><th>Nome</th><th>Tipo</th><th>Texto no Crachá</th><th>Escola</th><th>Status Geral</th><th>Hora</th><th>Abrir</th></tr>
            {linhas if linhas else '<tr><td colspan="8">Nenhum atleta cadastrado.</td></tr>'}
        </table>
        </div>
    </div>
    """
    return layout("Painel", conteudo, "painel")


# =========================
# RELATÓRIO COM FILTROS
# =========================

@app.route("/relatorio")
def relatorio():
    tipo_filtro = request.args.get("tipo", "").strip().upper()
    escola_filtro = request.args.get("escola", "").strip()
    categoria_filtro = request.args.get("categoria", "").strip()
    pessoa_filtro = request.args.get("pessoa", "").strip()
    busca = request.args.get("busca", "").strip()
    data_ini = request.args.get("data_ini", "").strip()
    data_fim = request.args.get("data_fim", "").strip()

    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    where = []
    params = []
    if tipo_filtro:
        where.append("c.tipo = ?"); params.append(tipo_filtro)
    if escola_filtro:
        where.append("a.escola LIKE ?"); params.append(f"%{escola_filtro}%")
    if categoria_filtro:
        where.append("a.categoria LIKE ?"); params.append(f"%{categoria_filtro}%")
    if pessoa_filtro:
        where.append("a.tipo_pessoa LIKE ?"); params.append(f"%{pessoa_filtro}%")
    if busca:
        where.append("(a.nome LIKE ? OR c.codigo_qr LIKE ?)"); params.extend([f"%{busca}%", f"%{busca}%"])
    if data_ini:
        where.append("substr(c.horario, 7, 4) || '-' || substr(c.horario, 4, 2) || '-' || substr(c.horario, 1, 2) >= ?"); params.append(data_ini)
    if data_fim:
        where.append("substr(c.horario, 7, 4) || '-' || substr(c.horario, 4, 2) || '-' || substr(c.horario, 1, 2) <= ?"); params.append(data_fim)

    where_sql = "WHERE " + " AND ".join(where) if where else ""

    cur.execute(f"""
        SELECT c.codigo_qr, c.tipo, c.horario, a.nome, a.tipo_pessoa, a.categoria, a.escola
        FROM checkins c
        LEFT JOIN atletas a ON a.codigo_qr = c.codigo_qr
        {where_sql}
        ORDER BY c.id DESC
    """, params)
    registros = cur.fetchall()

    cur.execute(f"""
        SELECT c.tipo, COUNT(*) as total
        FROM checkins c
        LEFT JOIN atletas a ON a.codigo_qr = c.codigo_qr
        {where_sql}
        GROUP BY c.tipo
        ORDER BY c.tipo
    """, params)
    resumo = cur.fetchall()
    conn.close()

    mapa = {r["tipo"]: r["total"] for r in resumo}
    total_filtrado = sum(mapa.values())

    cards = f'<div class="card"><div class="metric-label">Total Filtrado</div><div class="metric-value">{total_filtrado}</div></div>'
    for refeicao in REFEICOES_PADRAO:
        cards += f'<div class="card"><div class="metric-label">{h(refeicao)}</div><div class="metric-value">{mapa.get(refeicao, 0)}</div></div>'

    opcoes_tipo = '<option value="">TODOS</option>'
    for r in REFEICOES_PADRAO:
        selected = "selected" if tipo_filtro == r else ""
        opcoes_tipo += f'<option value="{h(r)}" {selected}>{h(r)}</option>'

    query = request.query_string.decode("utf-8")
    export_link = "/relatorio/exportar_csv" + ("?" + query if query else "")

    linhas = ""
    for r in registros:
        linhas += f"""
        <tr>
            <td>{h(r['horario'])}</td>
            <td><span class="pill green">{h(r['tipo'])}</span></td>
            <td>{h(r['codigo_qr'])}</td>
            <td>{h(r['nome'])}</td>
            <td>{h(r['tipo_pessoa'])}</td>
            <td>{h(r['categoria'])}</td>
            <td>{h(r['escola'])}</td>
        </tr>
        """

    conteudo = f"""
    <div class="topbar">
        <div><h1>Relatório de Refeições</h1><p>Filtre por refeição, escola, categoria, data, nome ou código.</p></div>
        <div class="actions"><a class="btn primary" href="/scanner">📷 Scanner</a><a class="btn blue" href="{export_link}">⬇️ Exportar CSV</a></div>
    </div>

    <form class="panel filters" method="GET" action="/relatorio" style="margin-top:0;">
        <div><label>Tipo</label><select name="tipo">{opcoes_tipo}</select></div>
        <div><label>Escola</label><input name="escola" value="{h(escola_filtro)}" placeholder="Ex: TOBIAS"></div>
        <div><label>Categoria/Texto</label><input name="categoria" value="{h(categoria_filtro)}" placeholder="Ex: 15 A 17 ou EEI"></div>
        <div><label>Tipo Pessoa</label><input name="pessoa" value="{h(pessoa_filtro)}" placeholder="Ex: ATLETA, TÉCNICO"></div>
        <div><label>Nome ou Código</label><input name="busca" value="{h(busca)}" placeholder="Nome ou QR"></div>
        <div><label>Data inicial</label><input type="date" name="data_ini" value="{h(data_ini)}"></div>
        <div><label>Data final</label><input type="date" name="data_fim" value="{h(data_fim)}"></div>
        <div><button class="btn primary" type="submit" style="width:100%;">FILTRAR</button></div>
        <div><a class="btn" href="/relatorio" style="width:100%;">LIMPAR</a></div>
    </form>

    <div class="grid">{cards}</div>

    <div class="panel">
        <h2>Registros encontrados: {len(registros)}</h2>
        <div class="table-wrap">
        <table>
            <tr><th>Horário</th><th>Tipo</th><th>Código QR</th><th>Nome</th><th>Tipo</th><th>Texto no Crachá</th><th>Escola</th></tr>
            {linhas if linhas else '<tr><td colspan="7">Nenhum registro encontrado.</td></tr>'}
        </table>
        </div>
    </div>
    """
    return layout("Relatórios", conteudo, "relatorio")


@app.route("/relatorio/exportar_csv")
def exportar_csv():
    tipo_filtro = request.args.get("tipo", "").strip().upper()
    escola_filtro = request.args.get("escola", "").strip()
    categoria_filtro = request.args.get("categoria", "").strip()
    pessoa_filtro = request.args.get("pessoa", "").strip()
    busca = request.args.get("busca", "").strip()
    data_ini = request.args.get("data_ini", "").strip()
    data_fim = request.args.get("data_fim", "").strip()

    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    where = []
    params = []
    if tipo_filtro:
        where.append("c.tipo = ?"); params.append(tipo_filtro)
    if escola_filtro:
        where.append("a.escola LIKE ?"); params.append(f"%{escola_filtro}%")
    if categoria_filtro:
        where.append("a.categoria LIKE ?"); params.append(f"%{categoria_filtro}%")
    if pessoa_filtro:
        where.append("a.tipo_pessoa LIKE ?"); params.append(f"%{pessoa_filtro}%")
    if busca:
        where.append("(a.nome LIKE ? OR c.codigo_qr LIKE ?)"); params.extend([f"%{busca}%", f"%{busca}%"])
    if data_ini:
        where.append("substr(c.horario, 7, 4) || '-' || substr(c.horario, 4, 2) || '-' || substr(c.horario, 1, 2) >= ?"); params.append(data_ini)
    if data_fim:
        where.append("substr(c.horario, 7, 4) || '-' || substr(c.horario, 4, 2) || '-' || substr(c.horario, 1, 2) <= ?"); params.append(data_fim)
    where_sql = "WHERE " + " AND ".join(where) if where else ""

    cur.execute(f"""
        SELECT c.horario, c.tipo, c.codigo_qr, a.nome, a.tipo_pessoa, a.categoria, a.escola
        FROM checkins c
        LEFT JOIN atletas a ON a.codigo_qr = c.codigo_qr
        {where_sql}
        ORDER BY c.id DESC
    """, params)
    registros = cur.fetchall()
    conn.close()

    output = io.StringIO()
    writer = csv.writer(output, delimiter=';')
    writer.writerow(["Horário", "Refeição/Controle", "Código QR", "Nome", "Tipo Pessoa", "Texto no Crachá", "Escola"])
    for r in registros:
        writer.writerow([r["horario"], r["tipo"], r["codigo_qr"], r["nome"] or "", r["tipo_pessoa"] or "", r["categoria"] or "", r["escola"] or ""])

    data = output.getvalue().encode("utf-8-sig")
    return send_file(io.BytesIO(data), as_attachment=True, download_name="relatorio_refeicoes.csv", mimetype="text/csv")


init_db()

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
