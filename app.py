import streamlit as st
import streamlit_authenticator as stauth
import pandas as pd
import mysql.connector
import os
from datetime import date, timedelta
from fpdf import FPDF

st.set_page_config(page_title="EARE — Sistema de Pedidos", layout="wide")

# ── AUTENTICAÇÃO ─────────────────────────────────────────────────────────────
# Usuários e senhas ficam no Streamlit Cloud: Settings > Secrets
# Perfis: "gestor" vê tudo | "multimarcas" só vê canal Multimarcas

credentials = {
    "usernames": {
        "raphael": {
            "name": "Raphael",
            "password": st.secrets["senha_raphael"],
        },
        "nath": {
            "name": "Nath",
            "password": st.secrets["senha_nath"],
        },
        "luca": {
            "name": "Luca",
            "password": st.secrets["senha_luca"],
        },
    }
}

PERFIS = {
    "raphael": "gestor",
    "nath":    "gestor",
    "luca":    "multimarcas",
}

authenticator = stauth.Authenticate(
    credentials,
    "eare_auth",          # nome do cookie
    "eare_chave_2026",    # chave secreta do cookie
    30,                   # dias para expirar
)

authenticator.login()

if st.session_state.get("authentication_status") is False:
    st.error("Usuário ou senha incorretos.")
    st.stop()

if st.session_state.get("authentication_status") is None:
    st.warning("Por favor, faça login para acessar o sistema.")
    st.stop()

# ── A partir daqui o usuário está autenticado ────────────────────────────────
usuario    = st.session_state["username"]
nome_user  = st.session_state["name"]
perfil     = PERFIS.get(usuario, "vendedor")
eh_gestor  = perfil == "gestor"

# ── LOGO E TÍTULO ────────────────────────────────────────────────────────────
if os.path.exists("logo.png"):
    col_logo, col_titulo = st.columns([1, 5])
    with col_logo:
        st.image("logo.png", width=180)
    with col_titulo:
        st.title("🪑 EARE — Sistema de Pedidos")
        st.caption(f"Olá, **{nome_user}** · Perfil: `{perfil}`")
else:
    st.title("🪑 EARE — Sistema de Pedidos")
    st.caption(f"Olá, **{nome_user}** · Perfil: `{perfil}`")

# Botão de logout na barra lateral
authenticator.logout("Sair", "sidebar")

# ── SENHA DO DRE (só pede se não for gestor) ──────────────────────────────────
try:
    SENHA_DRE = st.secrets["senha_dre"]
except Exception:
    SENHA_DRE = "eare2024"

# ── PRODUTOS VIA EXCEL ────────────────────────────────────────────────────────
PRODUTOS_FALLBACK = [
    {"area":"Interna","colecao":"Taquara","movel":"Mesa Centro","cor":"amber","bp":"Antares coffee table","desc":"mesa de centro","preco":3870,"custo":851.28},
    {"area":"Interna","colecao":"Taquara","movel":"Aparador","cor":"amber","bp":"Antares console table","desc":"aparador","preco":3110,"custo":596.47},
    {"area":"Interna","colecao":"Taquara","movel":"Mesa lateral","cor":"amber","bp":"Antares end table","desc":"mesa lateral","preco":1990,"custo":357.90},
    {"area":"Interna","colecao":"Guadua","movel":"Banco","cor":"caramelized","bp":"azara bench","desc":"banco caramelo","preco":2970,"custo":441.93},
    {"area":"Interna","colecao":"Guadua","movel":"Banco","cor":"sable","bp":"azara bench","desc":"banco café","preco":2970,"custo":441.93},
    {"area":"Interna","colecao":"Alyra","movel":"Poltrona","cor":"caramelized/red","bp":"Danica chair","desc":"poltrona caramelo/vermelho","preco":3540,"custo":577.95},
    {"area":"Interna","colecao":"Alyra","movel":"Poltrona","cor":"caramelized/grey","bp":"Danica chair","desc":"poltrona caramelo/cinza","preco":3540,"custo":425.46},
    {"area":"Interna","colecao":"Balcoa","movel":"Mesa de centro","cor":"amber","bp":"rosemary coffee table","desc":"mesa de centro","preco":4180,"custo":646.51},
    {"area":"Interna","colecao":"Balcoa","movel":"Mesa de centro","cor":"wheat","bp":"rosemary coffee table","desc":"mesa de centro","preco":4180,"custo":516.08},
    {"area":"Interna","colecao":"Balcoa","movel":"Mesa lateral","cor":"amber","bp":"thyme side table","desc":"mesa lateral","preco":1490,"custo":385.35},
    {"area":"Interna","colecao":"Balcoa","movel":"Mesa lateral","cor":"wheat","bp":"thyme side table","desc":"mesa lateral","preco":1490,"custo":363.51},
    {"area":"Interna","colecao":"Aurea","movel":"Mesa lateral","cor":"amber","bp":"sol side table","desc":"mesa lateral redonda","preco":1780,"custo":451.05},
    {"area":"Interna","colecao":"Aurea","movel":"Mesa lateral","cor":"wheat","bp":"sol side table","desc":"mesa lateral redonda","preco":1780,"custo":451.05},
    {"area":"Interna","colecao":"Aurea","movel":"Mesa de cabeceira","cor":"caramelized","bp":"currant nightstand","desc":"mesa de cabeceira","preco":3650,"custo":789.76},
    {"area":"Interna","colecao":"Aurea","movel":"Mesa de cabeceira","cor":"amber","bp":"currant nightstand","desc":"mesa de cabeceira","preco":3650,"custo":789.76},
    {"area":"Interna","colecao":"Aurea","movel":"Comoda","cor":"caramelized","bp":"currant dresser","desc":"comoda","preco":13460,"custo":3245.65},
    {"area":"Interna","colecao":"Aurea","movel":"Comoda","cor":"amber","bp":"currant dresser","desc":"comoda","preco":13890,"custo":3245.65},
    {"area":"Interna","colecao":"Aurea","movel":"Estante","cor":"caramelized","bp":"currant leaning bookshelf","desc":"estante de parede","preco":3590,"custo":541.14},
    {"area":"Interna","colecao":"Aurea","movel":"Estante","cor":"black walnut","bp":"currant leaning bookshelf","desc":"estante de parede","preco":3790,"custo":541.14},
    {"area":"Interna","colecao":"Kuruna","movel":"Mesa de cabeceira","cor":"wheat","bp":"monterey nightstand","desc":"mesa de cabeceira","preco":3290,"custo":813.28},
    {"area":"Interna","colecao":"Kuruna","movel":"Comoda","cor":"wheat","bp":"monterey dresser","desc":"comoda","preco":12220,"custo":3323.54},
    {"area":"Interna","colecao":"Valiha","movel":"Mesa de cabeceira","cor":"amber","bp":"ventura nightstand","desc":"mesa de cabeceira","preco":3305,"custo":813.28},
    {"area":"Interna","colecao":"Valiha","movel":"Comoda","cor":"amber","bp":"ventura dresser","desc":"comoda","preco":12220,"custo":3323.53},
    {"area":"Interna","colecao":"Arberella","movel":"Mesa de cabeceira","cor":"caramelized","bp":"willow nightstand","desc":"mesa de cabeceira","preco":3170,"custo":804.17},
    {"area":"Interna","colecao":"Arberella","movel":"Comoda","cor":"caramelized","bp":"willow dresser","desc":"comoda tradicional","preco":9110,"custo":2478.02},
    {"area":"Interna","colecao":"Sasa","movel":"Mesa de jantar","cor":"amber (erikka)","bp":"erikka dining table","desc":"mesa jantar variável","preco":14890,"custo":3399.10},
    {"area":"Interna","colecao":"Sasa","movel":"Mesa de jantar","cor":"amber (mija)","bp":"mija dining table","desc":"mesa jantar menor","preco":9980,"custo":1094.56},
    {"area":"Interna","colecao":"Sasa","movel":"Cadeira","cor":"amber","bp":"currant chair","desc":"cadeira clean","preco":1980,"custo":534.44},
    {"area":"Interna","colecao":"Sasa","movel":"Banqueta","cor":"caramelized","bp":"tulip bar stoll","desc":"banqueta curva","preco":1780,"custo":509.52},
    {"area":"Interna","colecao":"Pinga","movel":"Mesa de jantar","cor":"wheat","bp":"hanna dining table","desc":"mesa jantar orgânica 6 lug","preco":15990,"custo":1904.30},
    {"area":"Interna","colecao":"Pinga","movel":"Cadeira","cor":"wheat","bp":"hanna dining chair","desc":"cadeira design","preco":2680,"custo":697.57},
    {"area":"Interna","colecao":"Pinga","movel":"Aparador","cor":"wheat","bp":"hanna console/sideboard","desc":"aparador bambu","preco":12260,"custo":1126.52},
    {"area":"Interna","colecao":"Tulda","movel":"Mesa","cor":"wheat","bp":"sitka dining table","desc":"redonda 4 lugares","preco":4990,"custo":601.33},
    {"area":"Interna","colecao":"Tulda","movel":"Mesa","cor":"amber","bp":"sitka dining table","desc":"redonda 4 lugares","preco":5110,"custo":3106.62},
    {"area":"Interna","colecao":"Pariana","movel":"Banqueta","cor":"wheat","bp":"leif counter stoll","desc":"banqueta corda caramelo","preco":2110,"custo":645.38},
    {"area":"Interna","colecao":"Pariana","movel":"Banqueta","cor":"caviar","bp":"leif counter stoll","desc":"banqueta corda preta","preco":2110,"custo":645.38},
    {"area":"Externa","colecao":"Taquara","movel":"Mesa","cor":"—","bp":"Bled Table 210cm","desc":"mesa externa","preco":8110,"custo":895.29},
    {"area":"Externa","colecao":"Taquara","movel":"Cadeira","cor":"—","bp":"Bled Armchair","desc":"cadeira externa","preco":1940,"custo":527.40},
    {"area":"Externa","colecao":"Taquara","movel":"Espreguiçadeira","cor":"—","bp":"Bled Sun Lounger","desc":"espreguiçadeira","preco":5990,"custo":820.72},
    {"area":"Externa","colecao":"Taquara","movel":"Ombrelone","cor":"—","bp":"AA Sun Umbrella","desc":"ombrelone","preco":6530,"custo":1699.85},
    {"area":"Externa","colecao":"Pariana","movel":"Mesa","cor":"—","bp":"Positano Table 215cm","desc":"mesa externa","preco":8230,"custo":973.44},
    {"area":"Externa","colecao":"Pariana","movel":"Cadeira","cor":"—","bp":"Positano Chair","desc":"cadeira externa","preco":2600,"custo":742.54},
    {"area":"Externa","colecao":"Balcoa","movel":"Mesa redonda","cor":"—","bp":"Sunflower Table 98cm","desc":"mesa redonda","preco":5230,"custo":1147.60},
    {"area":"Externa","colecao":"Balcoa","movel":"Espreguiçadeira","cor":"—","bp":"Flexi Sun Lounger","desc":"espreguiçadeira","preco":11525,"custo":841.95},
    {"area":"Externa","colecao":"Alvimia","movel":"Sofá 3 lugares","cor":"—","bp":"Koloni 3-seater Sofa","desc":"sofá 3 lugares","preco":13000,"custo":2282.08},
    {"area":"Externa","colecao":"Alvimia","movel":"Sofá 1 lugar","cor":"—","bp":"Koloni Sofa","desc":"sofá 1 lugar","preco":5250,"custo":292.12},
    {"area":"Externa","colecao":"Alvimia","movel":"Mesa de centro","cor":"—","bp":"Koloni Table","desc":"mesa de centro","preco":5270,"custo":887.08},
    {"area":"Externa","colecao":"Nadinha","movel":"Cadeira","cor":"—","bp":"ET Chair","desc":"cadeira","preco":2650,"custo":689.46},
    {"area":"Externa","colecao":"Nadinha","movel":"Mesa","cor":"—","bp":"ET Table","desc":"mesa","preco":3120,"custo":695.83},
]

try:
    df_prod_global = pd.read_excel("produtos.xlsx")
    fonte_produtos = "📄 Excel"
except Exception:
    df_prod_global = pd.DataFrame(PRODUTOS_FALLBACK)
    fonte_produtos = "📋 Código (fallback)"

# ── TABELAS DE NEGÓCIO ────────────────────────────────────────────────────────
CANAIS_POR_PERFIL = {
    "gestor":      ["Direta", "Arquiteto", "Multimarcas"],
    "multimarcas": ["Multimarcas"],
}

DESC_REF = {
    "Direta":     {"À vista": 0.15, "A prazo": 0.07},
    "Arquiteto":  {"À vista": 0.10, "A prazo": 0.05},
    "Multimarcas":{"À vista": 0.05, "A prazo": 0.00},
}
COMISSAO = {"Direta": 0.0, "Arquiteto": 0.10, "Multimarcas": 0.40}
TAXAS_AVISTA = {"Pix (0%)": 0.0, "Débito (1,39%)": 0.0139, "Crédito à vista (3,49%)": 0.0349}

# ── BANCO DE DADOS (Supabase / PostgreSQL) ────────────────────────────────────
def get_conn():
    """Retorna conexão com MySQL Hostinger. Credenciais nos Secrets do Streamlit."""
    return mysql.connector.connect(
        host=st.secrets["db_host"],
        user=st.secrets["db_user"],
        password=st.secrets["db_password"],
        database=st.secrets["db_name"],
        charset="utf8mb4",
        use_unicode=True
    )

def init_db():
    """Garante que as tabelas existem — idempotente (pode rodar várias vezes)."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS pedidos (
            id INT AUTO_INCREMENT PRIMARY KEY,
            data TEXT, vendedor TEXT, cliente TEXT, projeto TEXT, validade TEXT,
            tipo_venda TEXT, forma_pagamento TEXT, parcelas INTEGER,
            subtotal REAL, desconto_pct REAL, receita REAL,
            cmv REAL, imposto REAL, taxa_pgto REAL, comissao REAL,
            frete REAL, montagem REAL, resultado REAL, margem REAL,
            itens TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS movimentos_estoque (
            id INT AUTO_INCREMENT PRIMARY KEY,
            data TEXT, pedido_id INTEGER,
            colecao TEXT, movel TEXT, cor TEXT,
            quantidade INTEGER,
            tipo TEXT
        )
    """)
    conn.commit()
    cur.close()
    conn.close()

def salvar_pedido(dados):
    """Insere pedido e retorna o ID gerado."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO pedidos (data,vendedor,cliente,projeto,validade,tipo_venda,
        forma_pagamento,parcelas,subtotal,desconto_pct,receita,cmv,imposto,
        taxa_pgto,comissao,frete,montagem,resultado,margem,itens)
        VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
    """, dados)
    pedido_id = cur.lastrowid
    conn.commit()
    cur.close()
    conn.close()
    return pedido_id

def carregar_pedidos(filtro_vendedor=None):
    conn = get_conn()
    cur = conn.cursor()
    if filtro_vendedor:
        cur.execute("SELECT * FROM pedidos WHERE vendedor=%s ORDER BY id DESC", (filtro_vendedor,))
    else:
        cur.execute("SELECT * FROM pedidos ORDER BY id DESC")
    rows = cur.fetchall()
    cols = [d[0] for d in cur.description] if cur.description else []
    cur.close()
    conn.close()
    return pd.DataFrame(rows, columns=cols)

# ── ESTOQUE ───────────────────────────────────────────────────────────────────
def calcular_estoque_produto(colecao, movel, cor, qtd_inicial):
    """Retorna estoque disponível: inicial - saídas + entradas."""
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("""
            SELECT
                COALESCE(SUM(CASE WHEN tipo='saida'  THEN quantidade ELSE 0 END), 0) as saidas,
                COALESCE(SUM(CASE WHEN tipo='entrada' THEN quantidade ELSE 0 END), 0) as entradas
            FROM movimentos_estoque
            WHERE colecao=%s AND movel=%s AND cor=%s
        """, (colecao, movel, cor))
        row = cur.fetchone()
        cur.close()
        conn.close()
        saidas, entradas = row if row else (0, 0)
        return max(0, int(qtd_inicial or 0) - int(saidas) + int(entradas))
    except Exception:
        return int(qtd_inicial or 0)

def registrar_saidas_estoque(pedido_id, cart):
    """Desconta itens vendidos do estoque."""
    conn = get_conn()
    cur = conn.cursor()
    for it in cart:
        cur.execute("""
            INSERT INTO movimentos_estoque (data, pedido_id, colecao, movel, cor, quantidade, tipo)
            VALUES (%s, %s, %s, %s, %s, %s, 'saida')
        """, (str(date.today()), pedido_id, it['colecao'], it['movel'], it['cor'], it['qtd']))
    conn.commit()
    cur.close()
    conn.close()

def registrar_entrada_estoque(colecao, movel, cor, quantidade):
    """Repõe estoque manualmente."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO movimentos_estoque (data, pedido_id, colecao, movel, cor, quantidade, tipo)
        VALUES (%s, NULL, %s, %s, %s, %s, 'entrada')
    """, (str(date.today()), colecao, movel, cor, quantidade))
    conn.commit()
    cur.close()
    conn.close()

def carregar_estoque_geral(df_prod):
    """Retorna DataFrame com estoque atual de todos os produtos."""
    if 'quantidade' not in df_prod.columns:
        return pd.DataFrame()
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        SELECT colecao, movel, cor,
            SUM(CASE WHEN tipo='saida'   THEN quantidade ELSE 0 END) as saidas,
            SUM(CASE WHEN tipo='entrada' THEN quantidade ELSE 0 END) as entradas
        FROM movimentos_estoque
        GROUP BY colecao, movel, cor
    """)
    rows = cur.fetchall()
    cur.close()
    conn.close()
    mov = pd.DataFrame(rows, columns=['colecao','movel','cor','saidas','entradas']) if rows else pd.DataFrame(columns=['colecao','movel','cor','saidas','entradas'])
    df = df_prod[['area','colecao','movel','cor','quantidade']].copy()
    df = df.merge(mov, on=['colecao','movel','cor'], how='left')
    df['saidas']   = df['saidas'].fillna(0).astype(int)
    df['entradas'] = df['entradas'].fillna(0).astype(int)
    df['disponivel'] = (df['quantidade'] - df['saidas'] + df['entradas']).clip(lower=0)
    df['status'] = df['disponivel'].apply(
        lambda x: '🔴 Esgotado' if x == 0 else ('🟡 Baixo' if x <= 2 else '🟢 OK')
    )
    return df[['area','colecao','movel','cor','quantidade','saidas','entradas','disponivel','status']]

# ── FUNÇÕES AUXILIARES ────────────────────────────────────────────────────────
def fmt(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def calcular(cart, tipo_venda, modalidade, meio_pgto, parcelas,
             desconto_pct, aliq_simples, frete, montagem):
    bruto    = sum(it["total"] for it in cart)
    cmv      = sum(it["totalCusto"] for it in cart)
    desc_val = bruto * desconto_pct
    receita  = bruto - desc_val
    v_simples = receita * aliq_simples
    canal     = COMISSAO[tipo_venda]
    v_canal   = receita * canal
    if modalidade == "À vista":
        taxa = TAXAS_AVISTA.get(meio_pgto, 0)
    else:
        taxa = 0.0279 + (parcelas - 1) * 0.0129
    v_taxa    = receita * taxa
    v_fm      = (frete + montagem) if bruto > 0 else 0
    resultado = receita - cmv - v_simples - v_taxa - v_canal - v_fm
    margem    = resultado / receita if receita > 0 else 0
    return dict(bruto=bruto, cmv=cmv, desc_val=desc_val, receita=receita,
                v_simples=v_simples, v_taxa=v_taxa, v_canal=v_canal,
                v_fm=v_fm, resultado=resultado, margem=margem, taxa=taxa)

# ── GERAÇÃO DE PDF ───────────────────────────────────────────────────────────
def gerar_pdf_orcamento(cart, nome_cliente, nome_projeto, validade_dt,
                         tipo_venda, modalidade, meio_pgto, parcelas,
                         desconto_pct, res, nome_vendedor):
    """Gera PDF do orçamento — apenas informações para o cliente, sem custos."""

    def s(text):
        """Garante compatibilidade Latin-1 para fontes Helvetica do FPDF."""
        return str(text).replace('—', '-').replace('–', '-').replace('’', "'")

    def fmt_pdf(v):
        return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=25)
    pdf.add_page()
    pdf.set_margins(20, 20, 20)

    # ── CABEÇALHO ──────────────────────────────────────────────────────────────
    logo_offset = 0
    if os.path.exists("logo.png"):
        pdf.image("logo.png", x=20, y=14, h=18)
        logo_offset = 45

    pdf.set_xy(20 + logo_offset, 14)
    pdf.set_font("Helvetica", "B", 15)
    pdf.set_text_color(40, 40, 40)
    pdf.cell(0, 7, "EARE", ln=True)

    pdf.set_x(20 + logo_offset)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(130, 130, 130)
    pdf.cell(0, 5, "Moveis em Bambu", ln=True)

    # Título e datas no canto direito
    pdf.set_xy(125, 14)
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(40, 40, 40)
    pdf.cell(65, 7, "ORCAMENTO", align="R", ln=True)

    pdf.set_xy(125, 21)
    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(130, 130, 130)
    pdf.cell(65, 4, f"Emitido: {date.today().strftime('%d/%m/%Y')}", align="R", ln=True)

    val_fmt = validade_dt.strftime("%d/%m/%Y") if hasattr(validade_dt, "strftime") else str(validade_dt)
    pdf.set_xy(125, 25)
    pdf.cell(65, 4, f"Valido ate: {val_fmt}", align="R", ln=True)

    # Linha divisória
    pdf.set_y(36)
    pdf.set_draw_color(200, 200, 200)
    pdf.set_line_width(0.4)
    pdf.line(20, pdf.get_y(), 190, pdf.get_y())
    pdf.ln(5)

    # ── DADOS DO CLIENTE ───────────────────────────────────────────────────────
    y0 = pdf.get_y()
    pdf.set_fill_color(247, 247, 247)
    pdf.rect(20, y0, 170, 20, "F")

    pdf.set_xy(25, y0 + 3)
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(28, 5, "CLIENTE")
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(40, 40, 40)
    pdf.cell(52, 5, s(nome_cliente))
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(28, 5, "PROJETO")
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(40, 40, 40)
    pdf.cell(0, 5, s(nome_projeto) if nome_projeto else "-", ln=True)

    pdf.set_xy(25, y0 + 11)
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(28, 5, "CANAL")
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(40, 40, 40)
    pdf.cell(52, 5, s(tipo_venda))
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(28, 5, "ATENDIMENTO")
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(40, 40, 40)
    pdf.cell(0, 5, s(nome_vendedor), ln=True)

    pdf.ln(7)

    # ── TABELA DE ITENS ────────────────────────────────────────────────────────
    # col_w soma 170 → tabela de x=20 a x=190 (margem igual dos dois lados)
    col_w   = [38, 34, 29, 11, 29, 29]
    headers = ["Colecao", "Movel", "Cor", "Qtd", "Preco Unit.", "Total"]
    aligns  = ["L", "L", "L", "C", "R", "R"]

    # Cabeçalho da tabela
    pdf.set_fill_color(40, 40, 40)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 8)
    for h, w, a in zip(headers, col_w, aligns):
        pdf.cell(w, 6, h, align=a, fill=True)
    pdf.ln()

    # Linhas dos produtos
    pdf.set_text_color(40, 40, 40)
    for idx, it in enumerate(cart):
        if idx % 2 == 0:
            pdf.set_fill_color(252, 252, 252)
        else:
            pdf.set_fill_color(255, 255, 255)
        pdf.set_font("Helvetica", "", 8)
        row = [s(it["colecao"]), s(it["movel"]), s(it["cor"]),
               str(it["qtd"]), fmt_pdf(it["preco"]), fmt_pdf(it["total"])]
        for val, w, a in zip(row, col_w, aligns):
            pdf.cell(w, 5.5, val, align=a, fill=True)
        pdf.ln()

    # Linha de fechamento da tabela
    pdf.set_draw_color(200, 200, 200)
    pdf.line(20, pdf.get_y(), 190, pdf.get_y())
    pdf.ln(2)

    # ── TOTAIS ─────────────────────────────────────────────────────────────────
    # Alinha com as colunas da tabela (x=20 a x=190):
    # cols 1-4 somam 112 → "Preco Unit." x=132→161, "Total" x=161→190
    tx = 20 + sum(col_w[:4])  # = 132
    lw = col_w[4]              # = 29 (largura "Preco Unit.")
    vw = col_w[5]              # = 29 (largura "Total")

    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(80, 80, 80)
    pdf.set_x(tx)
    pdf.cell(lw, 5, "Subtotal:", align="L")
    pdf.cell(vw, 5, fmt_pdf(res["bruto"]), align="R", ln=True)

    if desconto_pct > 0:
        pdf.set_x(tx)
        pdf.set_text_color(180, 50, 50)
        pdf.cell(lw, 5, f"Desconto ({desconto_pct*100:.1f}%):", align="L")
        pdf.cell(vw, 5, f"- {fmt_pdf(res['desc_val'])}", align="R", ln=True)
        pdf.set_text_color(80, 80, 80)

    pdf.set_draw_color(180, 180, 180)
    pdf.line(tx, pdf.get_y(), 190, pdf.get_y())
    pdf.ln(1)

    pdf.set_x(tx)
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(40, 40, 40)
    pdf.cell(lw, 6, "TOTAL:", align="L")
    pdf.cell(vw, 6, fmt_pdf(res["receita"]), align="R", ln=True)

    pdf.ln(5)

    # ── CONDIÇÕES DE PAGAMENTO ─────────────────────────────────────────────────
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(40, 40, 40)
    pdf.cell(0, 5, "Condicoes de Pagamento", ln=True)
    pdf.set_font("Helvetica", "", 9)
    pdf.set_text_color(80, 80, 80)

    if modalidade == "A prazo":
        pgto_txt = f"A prazo: {parcelas}x de {fmt_pdf(res['receita'] / parcelas)}"
    else:
        meio_pgto_label = meio_pgto.split("(")[0].strip()
        pgto_txt = f"A vista  -  {s(meio_pgto_label)}"

    pdf.cell(0, 5, pgto_txt, ln=True)

    pdf.ln(10)

    # ── RODAPÉ ─────────────────────────────────────────────────────────────────
    pdf.set_draw_color(200, 200, 200)
    pdf.line(20, pdf.get_y(), 190, pdf.get_y())
    pdf.ln(4)
    pdf.set_font("Helvetica", "I", 7)
    pdf.set_text_color(160, 160, 160)
    pdf.multi_cell(0, 4,
        f"Orcamento valido ate {val_fmt}. Valores sujeitos a alteracao apos esta data.\n"
        "Para duvidas ou confirmacao do pedido, entre em contato com nossa equipe.",
        align="C")

    return bytes(pdf.output())


# ── INICIALIZAÇÃO ─────────────────────────────────────────────────────────────
init_db()

if "cart" not in st.session_state:
    st.session_state.cart = []
if "senha_dre_ok" not in st.session_state:
    st.session_state.senha_dre_ok = eh_gestor  # gestor entra com DRE já aberto

# ── ABAS ──────────────────────────────────────────────────────────────────────
tabs_list = ["📋 Novo Pedido", "📦 Histórico de Pedidos"]
if eh_gestor:
    tabs_list.append("📊 Estoque")
aba = st.tabs(tabs_list)

with aba[0]:

    with st.expander("⚙️ Configurações", expanded=False):
        col1, col2, col3, col4 = st.columns(4)
        aliq_simples = col1.number_input("Simples Nacional (%)", value=9.0, step=0.5, min_value=0.0, max_value=20.0) / 100
        meta_ll      = col2.number_input("Meta LL (%)", value=25, step=1, min_value=0) / 100
        frete        = col3.number_input("Frete (R$)", value=200, step=10, min_value=0)
        montagem     = col4.number_input("Montagem (R$)", value=250, step=10, min_value=0)
        st.caption(f"Fonte dos produtos: {fonte_produtos}")

    st.subheader("👤 Cliente / Projeto")
    c1, c2, c3 = st.columns(3)
    nome_cliente = c1.text_input("Nome do cliente")
    nome_projeto = c2.text_input("Projeto / Referência")
    validade_map = {"7 dias": 7, "15 dias": 15, "30 dias": 30}
    validade_sel = c3.selectbox("Validade do orçamento", list(validade_map.keys()), index=1)
    validade_dt  = date.today() + timedelta(days=validade_map[validade_sel])

    st.subheader("🏷️ Tipo de Venda")
    canais_disponiveis = CANAIS_POR_PERFIL.get(perfil, ["Direta"])
    tipo_venda = st.radio("", canais_disponiveis, horizontal=True,
                          help="Direta: sem comissão | Arquiteto: 10% | Multimarcas: 40%")
    if tipo_venda == "Arquiteto":
        st.caption("💡 Comissão de 10% incluída no cálculo do resultado.")
    elif tipo_venda == "Multimarcas":
        st.caption("💡 Comissão de 40% incluída. Verifique a margem antes de dar desconto.")

    st.subheader("🔍 Busca de Produto")
    df_prod = df_prod_global.copy()

    b1, b2, b3, b4 = st.columns(4)
    areas    = [""] + sorted(df_prod["area"].unique().tolist())
    sel_area = b1.selectbox("Área", areas)

    colecoes = [""]
    if sel_area:
        colecoes += sorted(df_prod[df_prod["area"] == sel_area]["colecao"].unique().tolist())
    sel_col = b2.selectbox("Coleção (EARE)", colecoes, disabled=not sel_area)

    moveis = [""]
    if sel_col:
        moveis += sorted(df_prod[(df_prod["area"] == sel_area) & (df_prod["colecao"] == sel_col)]["movel"].unique().tolist())
    sel_mov = b3.selectbox("Móvel", moveis, disabled=not sel_col)

    cores = [""]
    if sel_mov:
        cores += sorted(df_prod[(df_prod["area"] == sel_area) & (df_prod["colecao"] == sel_col) & (df_prod["movel"] == sel_mov)]["cor"].unique().tolist())
    sel_cor = b4.selectbox("Cor", cores, disabled=not sel_mov)

    prod_sel = None
    if sel_cor:
        mask = ((df_prod["area"] == sel_area) & (df_prod["colecao"] == sel_col) &
                (df_prod["movel"] == sel_mov) & (df_prod["cor"] == sel_cor))
        prod_sel = df_prod[mask].iloc[0].to_dict() if mask.any() else None

    if prod_sel:
        # ── Verifica estoque ──────────────────────────────────────────────────
        tem_estoque = 'quantidade' in df_prod.columns
        estoque_atual = calcular_estoque_produto(
            prod_sel['colecao'], prod_sel['movel'], prod_sel['cor'],
            prod_sel.get('quantidade', 999)
        ) if tem_estoque else 999

        i1, i2, i3 = st.columns([2, 2, 1])
        i1.text_input("Fornecedor (BP)", value=prod_sel["bp"], disabled=True)
        i2.text_input("Descrição", value=prod_sel["desc"], disabled=True)
        qtd_max = max(1, estoque_atual)
        qtd = i3.number_input("Qtd", min_value=1, max_value=qtd_max if tem_estoque else 999, value=1)

        col_preco, col_est = st.columns([3, 1])
        col_preco.metric("Preço final (VL_VENDA_FINAL)", fmt(prod_sel["preco"]))
        if tem_estoque:
            if estoque_atual <= 0:
                col_est.error("Sem estoque")
            elif estoque_atual <= 2:
                col_est.warning(f"Estoque: {estoque_atual}")
            else:
                col_est.success(f"Estoque: {estoque_atual}")

        sem_estoque = tem_estoque and estoque_atual <= 0
        if sem_estoque:
            st.error("⚠️ Produto esgotado — não é possível adicionar ao pedido.")
        elif st.button("➕ Adicionar ao pedido", type="primary"):
            # Checa se a qtd no carrinho + nova qtd não ultrapassa o estoque
            no_cart = sum(it['qtd'] for it in st.session_state.cart
                          if it['colecao'] == prod_sel['colecao']
                          and it['movel'] == prod_sel['movel']
                          and it['cor'] == prod_sel['cor'])
            if tem_estoque and no_cart + qtd > estoque_atual:
                st.error(f"Quantidade indisponível. Estoque livre: {estoque_atual - no_cart}")
            else:
                st.session_state.cart.append({
                    "colecao":    prod_sel["colecao"],
                    "movel":      prod_sel["movel"],
                    "cor":        prod_sel["cor"],
                    "qtd":        qtd,
                    "preco":      prod_sel["preco"],
                    "custo":      prod_sel["custo"],
                    "total":      prod_sel["preco"] * qtd,
                    "totalCusto": prod_sel["custo"] * qtd,
                })
                st.success(f"✅ {prod_sel['movel']} ({prod_sel['cor']}) adicionado!")
                st.rerun()

    st.subheader("🛒 Itens do Pedido")
    if not st.session_state.cart:
        st.info("Nenhum item adicionado ainda.")
    else:
        for i, it in enumerate(st.session_state.cart):
            c1, c2, c3, c4, c5, c6 = st.columns([2, 2, 2, 1, 2, 1])
            c1.write(it["colecao"])
            c2.write(it["movel"])
            c3.write(it["cor"])
            c4.write(f"{it['qtd']}x")
            c5.write(fmt(it["total"]))
            if c6.button("🗑️", key=f"del_{i}"):
                st.session_state.cart.pop(i)
                st.rerun()
        total_qtd   = sum(it["qtd"] for it in st.session_state.cart)
        total_bruto = sum(it["total"] for it in st.session_state.cart)
        st.markdown(f"**Total: {total_qtd} itens — {fmt(total_bruto)}**")
        if st.button("🗑️ Limpar carrinho"):
            st.session_state.cart = []
            st.rerun()

    st.subheader("💸 Desconto")
    d1, d2, d3 = st.columns(3)
    bruto_atual = sum(it["total"] for it in st.session_state.cart)
    tipo_desconto = d1.radio("Tipo de desconto", ["%", "R$"], horizontal=True)
    if tipo_desconto == "%":
        desc_input = d1.number_input("Desconto (%)", min_value=0.0, max_value=100.0, step=0.5, value=0.0)
        desconto_pct = desc_input / 100
    else:
        desc_val_rs = d1.number_input("Desconto (R$)", min_value=0.0, max_value=float(bruto_atual) if bruto_atual > 0 else 0.0, step=10.0, value=0.0)
        desconto_pct = desc_val_rs / bruto_atual if bruto_atual > 0 else 0.0
    modalidade   = d2.radio("Modalidade", ["À vista", "A prazo"], horizontal=True)
    ref_key      = "À vista" if modalidade == "À vista" else "A prazo"
    ref_desc     = DESC_REF[tipo_venda][ref_key]
    if desconto_pct > ref_desc:
        d2.warning(f"⚠️ Acima da referência ({ref_desc*100:.0f}% máx)")
    else:
        d2.success(f"✅ Dentro da referência ({ref_desc*100:.0f}% máx)")
    d3.metric("Subtotal com desconto", fmt(bruto_atual * (1 - desconto_pct)))

    st.subheader("💳 Forma de Pagamento")
    p1, p2 = st.columns(2)
    meio_pgto = ""
    parcelas  = 1
    if modalidade == "À vista":
        meio_pgto = p1.radio("Meio de pagamento", list(TAXAS_AVISTA.keys()), horizontal=True)
    else:
        parcelas  = p1.selectbox("Nº de parcelas", list(range(2, 11)),
                                  format_func=lambda n: f"{n}x — taxa {(0.0279+(n-1)*0.0129)*100:.2f}%")
        taxa_calc = 0.0279 + (parcelas - 1) * 0.0129
        p2.info(f"Taxa de {taxa_calc*100:.2f}% (14 dias corridos)")

    # ── DRE ───────────────────────────────────────────────────────────────────
    if st.session_state.cart:
        res = calcular(st.session_state.cart, tipo_venda, modalidade, meio_pgto,
                       parcelas, desconto_pct, aliq_simples, frete, montagem)

        st.subheader("📊 DRE do Pedido")

        if not st.session_state.senha_dre_ok:
            with st.container(border=True):
                st.markdown("🔒 **Área restrita — Gestão**")
                st.caption("O DRE completo é visível apenas para gestores.")
                col_inp, col_btn, _ = st.columns([2, 1, 3])
                senha_input = col_inp.text_input("Senha", type="password",
                                                  label_visibility="collapsed",
                                                  placeholder="Digite a senha do DRE")
                if col_btn.button("Acessar", type="primary"):
                    if senha_input == SENHA_DRE:
                        st.session_state.senha_dre_ok = True
                        st.rerun()
                    else:
                        st.error("Senha incorreta.")
        else:
            col_dre, col_alert = st.columns([3, 2])
            with col_dre:
                st.markdown(f"**(+) Receita bruta (tabela):** {fmt(res['bruto'])}")
                if res['desc_val'] > 0:
                    st.markdown(f"**(−) Desconto:** ({fmt(res['desc_val'])})")
                st.markdown(f"**(=) Receita reconhecida:** **{fmt(res['receita'])}**")
                st.divider()
                pct = lambda v: f"{(v/res['receita']*100):.1f}%" if res['receita'] > 0 else "—"
                st.markdown(f"(−) CMV — custo NF: ({fmt(res['cmv'])}) `{pct(res['cmv'])}`")
                st.markdown(f"(−) Simples Nacional ({aliq_simples*100:.1f}%): ({fmt(res['v_simples'])}) `{pct(res['v_simples'])}`")
                st.markdown(f"(−) Taxa de pagamento ({res['taxa']*100:.2f}%): ({fmt(res['v_taxa'])}) `{pct(res['v_taxa'])}`")
                if res['v_canal'] > 0:
                    st.markdown(f"(−) Comissão {tipo_venda}: ({fmt(res['v_canal'])}) `{pct(res['v_canal'])}`")
                st.markdown(f"(−) Frete + Montagem: ({fmt(res['v_fm'])}) `{pct(res['v_fm'])}`")
                st.divider()
                mg_color = "normal" if res['margem'] >= meta_ll else "inverse"
                st.metric("(=) Resultado Operacional",
                          fmt(res['resultado']),
                          f"{res['margem']*100:.1f}% (meta: {meta_ll*100:.0f}%)",
                          delta_color=mg_color)
                st.progress(min(1.0, max(0.0, res['margem'] / meta_ll if meta_ll > 0 else 0)))
                if modalidade == "A prazo":
                    st.info(f"💳 Parcela do cliente: **{fmt(res['receita']/parcelas)}/mês** ({parcelas}x)")
                if not eh_gestor:
                    if st.button("🔒 Fechar DRE"):
                        st.session_state.senha_dre_ok = False
                        st.rerun()

            with col_alert:
                st.markdown("**🚦 Alertas**")
                alertas = []
                if res['resultado'] < 0:
                    alertas.append(("🔴", "PREJUÍZO. Não feche sem aprovação do gestor."))
                elif res['margem'] < meta_ll * 0.5:
                    alertas.append(("🔴", f"Margem crítica ({res['margem']*100:.1f}%). Aprovação necessária."))
                elif res['margem'] < meta_ll:
                    alertas.append(("🟡", f"Margem {res['margem']*100:.1f}% abaixo da meta ({meta_ll*100:.0f}%). Revise."))
                else:
                    alertas.append(("🟢", f"Margem {res['margem']*100:.1f}% — dentro do target!"))
                if desconto_pct > ref_desc:
                    alertas.append(("🟡", f"Desconto {desconto_pct*100:.1f}% acima da referência {ref_desc*100:.0f}%."))
                if modalidade == "A prazo" and parcelas >= 8:
                    alertas.append(("🟡", f"Parcelado em {parcelas}x: taxa {res['taxa']*100:.2f}%. Incentive Pix."))
                if tipo_venda == "Multimarcas":
                    alertas.append(("💡", f"Comissão 40% = {fmt(res['v_canal'])}. Verifique a margem antes do desconto."))
                elif tipo_venda == "Arquiteto":
                    alertas.append(("💡", f"Comissão 10% = {fmt(res['v_canal'])}. Alinhe prazo do pagamento."))
                if res['receita'] > 0 and res['v_fm'] / res['receita'] > 0.10:
                    alertas.append(("🟡", f"Frete+montagem = {res['v_fm']/res['receita']*100:.1f}% da receita."))
                for icon, msg in alertas:
                    st.markdown(f"{icon} {msg}")

        st.divider()
        st.subheader("💾 Salvar Pedido")

        # ── Botão de PDF ──────────────────────────────────────────────────────
        if nome_cliente:
            pdf_bytes = gerar_pdf_orcamento(
                cart=st.session_state.cart,
                nome_cliente=nome_cliente,
                nome_projeto=nome_projeto,
                validade_dt=validade_dt,
                tipo_venda=tipo_venda,
                modalidade=modalidade,
                meio_pgto=meio_pgto,
                parcelas=parcelas,
                desconto_pct=desconto_pct,
                res=res,
                nome_vendedor=nome_user,
            )
            nome_arquivo = f"Orcamento_EARE_{nome_cliente.replace(' ', '_')}_{date.today()}.pdf"
            st.download_button(
                label="📄 Baixar Orçamento PDF",
                data=pdf_bytes,
                file_name=nome_arquivo,
                mime="application/pdf",
            )
        else:
            st.info("Preencha o nome do cliente para gerar o PDF.")

        if st.button("✅ Confirmar e Salvar Pedido", type="primary"):
            if not nome_cliente:
                st.error("Informe o nome do cliente antes de salvar.")
            else:
                import json
                itens_json = json.dumps(st.session_state.cart)
                pedido_id = salvar_pedido((
                    str(date.today()), nome_user, nome_cliente, nome_projeto,
                    str(validade_dt), tipo_venda,
                    meio_pgto if modalidade == "À vista" else f"{parcelas}x",
                    parcelas if modalidade == "A prazo" else 1,
                    res['bruto'], desconto_pct * 100, res['receita'],
                    res['cmv'], res['v_simples'], res['v_taxa'],
                    res['v_canal'], frete, montagem,
                    res['resultado'], res['margem'] * 100, itens_json
                ))
                # Desconta automaticamente do estoque
                registrar_saidas_estoque(pedido_id, st.session_state.cart)
                st.success(f"✅ Pedido de {nome_cliente} salvo e estoque atualizado!")
                st.session_state.cart = []
                st.rerun()

with aba[1]:
    st.subheader("📦 Histórico de Pedidos")


    # Gestor vê todos; outros perfis veem só os próprios pedidos
    if eh_gestor:
        df_hist = carregar_pedidos()
    else:
        df_hist = carregar_pedidos(filtro_vendedor=nome_user)

    if df_hist.empty:
        st.info("Nenhum pedido salvo ainda.")
    else:
        colunas = ["id", "data", "vendedor", "cliente", "projeto",
                   "tipo_venda", "forma_pagamento", "receita", "desconto_pct",
                   "resultado", "margem"]
        colunas_existentes = [c for c in colunas if c in df_hist.columns]
        df_show = df_hist[colunas_existentes].copy()
        df_show.rename(columns={
            "id": "#", "data": "Data", "vendedor": "Vendedor",
            "cliente": "Cliente", "projeto": "Projeto",
            "tipo_venda": "Canal", "forma_pagamento": "Pagamento",
            "receita": "Receita (R$)", "desconto_pct": "Desc %",
            "resultado": "Resultado (R$)", "margem": "Margem %"
        }, inplace=True)
        if "Receita (R$)"   in df_show.columns: df_show["Receita (R$)"]   = df_show["Receita (R$)"].map(lambda x: f"R$ {x:,.0f}")
        if "Resultado (R$)" in df_show.columns: df_show["Resultado (R$)"] = df_show["Resultado (R$)"].map(lambda x: f"R$ {x:,.0f}")
        if "Margem %"       in df_show.columns: df_show["Margem %"]       = df_show["Margem %"].map(lambda x: f"{x:.1f}%")
        if "Desc %"         in df_show.columns: df_show["Desc %"]         = df_show["Desc %"].map(lambda x: f"{x:.1f}%")
        st.dataframe(df_show, use_container_width=True)
        st.metric("Total de pedidos", len(df_hist))

# ── ABA ESTOQUE (só para gestores) ───────────────────────────────────────────
if eh_gestor and len(aba) > 2:
    with aba[2]:
        st.subheader("📊 Controle de Estoque")

        if 'quantidade' not in df_prod_global.columns:
            st.warning("Adicione a coluna **quantidade** no arquivo `produtos.xlsx` para ativar o estoque.")
            st.caption("Exemplo: coluna 'quantidade' com o número de unidades disponíveis para cada produto.")
        else:
            df_est = carregar_estoque_geral(df_prod_global)

            # Métricas resumo
            m1, m2, m3 = st.columns(3)
            m1.metric("Total de produtos", len(df_est))
            m2.metric("Itens disponíveis", int(df_est['disponivel'].sum()))
            m3.metric("Produtos esgotados", int((df_est['disponivel'] == 0).sum()))

            st.divider()

            # Filtro de status
            filtro_status = st.radio("Filtrar por status", ["Todos", "🟢 OK", "🟡 Baixo", "🔴 Esgotado"], horizontal=True)
            df_view = df_est if filtro_status == "Todos" else df_est[df_est['status'] == filtro_status]

            df_view = df_view.rename(columns={
                'area': 'Área', 'colecao': 'Coleção', 'movel': 'Móvel', 'cor': 'Cor',
                'quantidade': 'Inicial', 'saidas': 'Saídas', 'entradas': 'Entradas',
                'disponivel': 'Disponível', 'status': 'Status'
            })
            st.dataframe(df_view, use_container_width=True)

            st.divider()

            # Formulário de reposição
            st.subheader("➕ Repor Estoque")
            with st.form("repor_estoque"):
                r1, r2, r3, r4 = st.columns(4)
                opcoes_col = [""] + sorted(df_prod_global["colecao"].unique().tolist())
                rep_col = r1.selectbox("Coleção", opcoes_col)
                opcoes_mov = [""] + sorted(df_prod_global[df_prod_global["colecao"] == rep_col]["movel"].unique().tolist()) if rep_col else [""]
                rep_mov = r2.selectbox("Móvel", opcoes_mov)
                opcoes_cor = [""] + sorted(df_prod_global[(df_prod_global["colecao"] == rep_col) & (df_prod_global["movel"] == rep_mov)]["cor"].unique().tolist()) if rep_mov else [""]
                rep_cor = r3.selectbox("Cor", opcoes_cor)
                rep_qtd = r4.number_input("Qtd a repor", min_value=1, value=1)
                if st.form_submit_button("✅ Confirmar reposição", type="primary"):
                    if rep_col and rep_mov and rep_cor:
                        registrar_entrada_estoque(rep_col, rep_mov, rep_cor, rep_qtd)
                        st.success(f"✅ {rep_qtd} unidade(s) de {rep_mov} ({rep_cor}) repostas!")
                        st.rerun()
                    else:
                        st.error("Selecione coleção, móvel e cor.")
