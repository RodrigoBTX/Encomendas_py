from flask import Flask, render_template, request, jsonify,redirect, url_for, session, make_response
import pyodbc
import pandas as pd
import configparser
import os
import base64
import webbrowser
from threading import Timer



app = Flask(__name__)
# app = Flask(__name__, template_folder="templates")
app.secret_key = os.urandom(24)  # Chave segura para sessões

def ler_versao_local():
    if os.path.exists("version.txt"):
        with open("version.txt", "r") as f:
            return f.read().strip()
    return "0.0.0"

# Pasta fixa para configs
CONFIG_DIR = os.path.join(os.path.expanduser("~"), "LacovianaConfig")
if not os.path.exists(CONFIG_DIR):
    os.makedirs(CONFIG_DIR)

CONFIG_FILE = os.path.join(CONFIG_DIR, "config.ini")


# Guardar configs

# CONFIG_FILE = "config.ini"

def ler_config():
    config = configparser.ConfigParser()
    if os.path.exists(CONFIG_FILE):
        config.read(CONFIG_FILE)
        dsn = config.get("DATABASE", "dsn", fallback="Lacoviana")
        db = config.get("DATABASE", "database", fallback="")
        user = config.get("DATABASE", "user", fallback="")
        password_b64 = config.get("DATABASE", "password", fallback="")
        # Decodifica a password de base64
        password = base64.b64decode(password_b64.encode('utf-8')).decode('utf-8') if password_b64 else ""
        return dsn, db, user, password
    else:
        return "Lacoviana", None, None, None


def salvar_config(dsn,database, user, password):
    config = configparser.ConfigParser()
    password_b64 = base64.b64encode(password.encode('utf-8')).decode('utf-8')
    config["DATABASE"] = {
        "dsn": dsn,
        "database": database,
        "user": user,
        "password": password_b64
    }
    with open(CONFIG_FILE, "w") as f:
        config.write(f)


# Usa o DSN criado no Windows
# conn_str = "DSN=Lacoviana;UID=admin-btx;PWD=;DATABASE=PHC_20240310"
def criar_conexao():
    dsn, db, user, password = ler_config()
    if not all([dsn,db, user, password]):
        return None  # sinaliza que configuração não existe
    # DSN fixo, mas database, user e password dinâmicos
    conn_str = f"DSN={dsn};UID={user};PWD={password};DATABASE={db}"
    return pyodbc.connect(conn_str)


def get_encomendas(filtros):
    try:
        #conn = pyodbc.connect(conn_str)
        #cursor = conn.cursor()

        conn = criar_conexao()
        if conn is None:
            return pd.DataFrame(), "Configuração da BD ausente. É necessário configurar a ligação á Base de dados."

        cursor = conn.cursor()

    # teste SP

    # print("Parâmetros enviados para SP:", (
    # filtros.get("data_ini"),
    # filtros.get("data_fin"),
    # filtros.get("cliente_ini"),
    # filtros.get("cliente_fin"),
    # filtros.get("trat_ini"),
    # filtros.get("trat_fin"),
    # filtros.get("requisicao"),
    # filtros.get("enc_ini"),
    # filtros.get("enc_fin"),
    # filtros.get("tipo"),
    # filtros.get("subtipo"),
    # filtros.get("gama_cor"),
    # filtros.get("linha"),
    # filtros.get("ordem")
    #     )
    # )


        # Chamar o SP passando os parâmetros
        cursor.execute("""
        EXEC sp_ListarEncomendas 
            @data_ini=?, @data_fin=?, 
            @cliente_ini=?, @cliente_fin=?, 
            @trat_ini=?, @trat_fin=?, 
            @req=?, 
            @enc_ini=?, @enc_fin=?, 
            @tipo=?, @subtipo=?, 
            @gamacor=?, @linha=?, @ordem=?
        """, (
            filtros.get("data_ini"),
            filtros.get("data_fin"),
            filtros.get("cliente_ini"),
            filtros.get("cliente_fin"),
            filtros.get("trat_ini"),
            filtros.get("trat_fin"),
            filtros.get("requisicao"),
            filtros.get("enc_ini"),
            filtros.get("enc_fin"),
            filtros.get("tipo"),
            filtros.get("subtipo"),
            filtros.get("gama_cor"),
            filtros.get("linha"),
            filtros.get("ordem")
            )
        )

        rows = cursor.fetchall()
        columns = [desc[0] for desc in cursor.description] if rows else []

        cursor.close()
        conn.close()

        # Devolve em DataFrame
        df = pd.DataFrame.from_records(rows, columns=columns) if rows else pd.DataFrame()
        return df, None

    except pyodbc.Error as e:
        return pd.DataFrame(), f"Erro ao consultar a base de dados: {e}"


## Autenticação básica para a página de configurações

# ------------------------------
# Login via formulário
# ------------------------------
ADMIN_USER = "sa"
ADMIN_PASSWORD = "admin-btx"

@app.route("/login", methods=["GET", "POST"])
def login():
    error_msg = None
    if request.method == "POST":
        user = request.form.get("user")
        password = request.form.get("password")
        if user == ADMIN_USER and password == ADMIN_PASSWORD:
            #session["admin_logged"] = True
            return redirect(url_for("configuracoes"))
        else:
            error_msg = "User ou senha incorretos."
    return render_template("login.html", error_msg=error_msg)

#@app.route("/logout")
#def logout():
#    session.pop("admin_logged", None)
#    return redirect(url_for("index"))



# ------------------------------
# Página de configurações
# ------------------------------
@app.route("/configuracoes", methods=["GET", "POST"])
def configuracoes():
    #if not session.get("admin_logged"):
     #   return redirect(url_for("login"))
    
    dsn,db, user, password = ler_config()
    error_msg = None  # Para mensagens de erro

    if request.method == "POST":
        dsn_form = request.form.get("dsn")
        database_form = request.form.get("database")
        user_form = request.form.get("user")
        password_form = request.form.get("password")

        #salvar_config(database, user, password)
        #return redirect(url_for("index"))
    
        # Testar a conexão antes de salvar
        try:
            test_conn_str = f"DSN={dsn_form};UID={user_form};PWD={password_form};DATABASE={database_form}"
            conn = pyodbc.connect(test_conn_str)
            conn.close()
            # Se não der erro, salva no config.ini
            salvar_config(dsn_form,database_form, user_form, password_form)
            return redirect(url_for("index"))
        except pyodbc.Error as e:
            error_msg = f"Erro ao conectar na base de dados: {e}"
        except Exception as e:
            error_msg = f"Ocorreu um erro inesperado: {e}"

            # Se der erro, mantém os valores digitados para corrigir
        dsn, db, user, password = database_form, user_form, password_form

    return render_template(
        "configuracoes.html",
        dsn=dsn or "Lacoviana",
        database=db or "",
        user=user or "",
        password=password or "",
        error_msg=error_msg
    )


@app.route("/tratamentos")
def tratamentos():
    #conn = pyodbc.connect(conn_str)
    conn = criar_conexao()
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT tratamento FROM u_tratamentos (NOLOCK) order by tratamento asc")
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return jsonify([r[0] for r in rows])


@app.route("/clientes")
def clientes():
    #conn = pyodbc.connect(conn_str)
    conn = criar_conexao()
    cursor = conn.cursor()
    cursor.execute("SELECT nome FROM cl (NOLOCK) order by nome asc")
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    return jsonify([r[0] for r in rows])


@app.route("/linhas")
def linhas():
    #conn = pyodbc.connect(conn_str)
    conn = criar_conexao()
    cursor = conn.cursor()
    cursor.execute("SELECT DISTINCT linha, linha FROM u_tratamentos (NOLOCK)")
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    # devolve lista de tuplas (valor, display)
    return jsonify([{"value": r[0], "label": r[1]} for r in rows])



# executar sps para a impressão
def executar_sps(filtros):
    # filtros = request.form.to_dict()

    try:
        conn = criar_conexao()
        cursor = conn.cursor()

        # 1. Clientes
        cursor.execute("""
            EXEC sp_clientes 
                @data_ini=?, @data_fin=?, 
                @cliente_ini=?, @cliente_fin=?, 
                @trat_ini=?, @trat_fin=?, 
                @req=?, 
                @enc_ini=?, @enc_fin=?, 
                @tipo=?, @subtipo=?, 
                @gamacor=?, @linha=?, @ordem=? 
        """, (
            filtros.get("data_ini"),
            filtros.get("data_fin"),
            filtros.get("cliente_ini"),
            filtros.get("cliente_fin"),
            filtros.get("trat_ini"),
            filtros.get("trat_fin"),
            filtros.get("requisicao"),
            filtros.get("enc_ini"),
            filtros.get("enc_fin"),
            filtros.get("tipo"),
            filtros.get("subtipo"),
            filtros.get("gama_cor"),
            filtros.get("linha"),
            filtros.get("ordem") or 1
        ))
        clientes = cursor.fetchall()
        clientes_cols = [c[0] for c in cursor.description]

        resultado = []
        for cliente in clientes:
            cliente_dict = dict(zip(clientes_cols, cliente))
            cliente_nome = cliente_dict.get("cliente")
            tratamento_cliente = cliente_dict.get("tratamento")

            # 2. Encomendas desse cliente
            cursor.execute("""
                EXEC sp_encomendas 
                    @data_ini=?, @data_fin=?, 
                    @req=?, 
                    @enc_ini=?, @enc_fin=?, 
                    @tipo=?, @subtipo=?, 
                    @gamacor=?, @linha=?, 
                    @cliente=?, @tratamento=? 
            """, (
                filtros.get("data_ini"),
                filtros.get("data_fin"),
                filtros.get("requisicao"),
                filtros.get("enc_ini"),
                filtros.get("enc_fin"),
                filtros.get("tipo"),
                filtros.get("subtipo"),
                filtros.get("gama_cor"),
                filtros.get("linha"),
                cliente_nome,
                tratamento_cliente
            ))
            encomendas = cursor.fetchall()
            encomendas_cols = [c[0] for c in cursor.description]

            encomendas_data = []
            for enc in encomendas:
                enc_dict = dict(zip(encomendas_cols, enc))
                 # Atenção: o SP encomendas já devolve obrano e obranome
                obrano_enc = enc_dict.get("obrano")       # número da encomenda
                obranome_enc = enc_dict.get("obranome")   # nome da obra -> vai no @req
                trat_enc = enc_dict.get("tratamento")
                cliente_nome = cliente_dict.get("cliente")

                # 3. Linhas da encomenda
                cursor.execute("""
                    EXEC sp_linhas 
                        @req=?, 
                        @enc=?, 
                        @tipo=?, @subtipo=?, 
                        @gamacor=?, @linha=?, 
                        @cliente=?, @tratamento=? 
                """, (
                    obranome_enc,
                    obrano_enc,
                    filtros.get("tipo"),
                    filtros.get("subtipo"),
                    filtros.get("gama_cor"),
                    filtros.get("linha"),
                    cliente_nome,
                    trat_enc
                ))
                linhas = cursor.fetchall()
                linhas_cols = [c[0] for c in cursor.description]

                linhas_data = [dict(zip(linhas_cols, l)) for l in linhas]

                encomendas_data.append({
                    "dados": enc_dict,
                    "linhas": linhas_data
                })

            resultado.append({
                "cliente": cliente_dict,
                "encomendas": encomendas_data
            })

        cursor.close()
        conn.close()
        return resultado
    
    except Exception as e:
        raise(Exception(f"Erro ao executar SPs: {e}"))


# impressao em PDF

from flask import make_response, request
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether, HRFlowable
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from datetime import datetime
import io
from reportlab.pdfgen import canvas as pdf_canvas

# Canvas personalizado para numerar páginas "Página X de Y"
class NumberedCanvas(pdf_canvas.Canvas):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            super().showPage()
        super().save()

    def draw_page_number(self, page_count):
        largura, _ = A4
        self.setFont("Helvetica", 8)
        self.drawCentredString(largura / 2, 20, f"Página {self._pageNumber} de {page_count}")

def cabecalho(canvas, doc, filtros):
    largura, altura = A4
    agora = datetime.now().strftime("%d-%m-%Y %H:%M")

    canvas.setLineWidth(0.8)
    canvas.line(30, altura - 40, largura - 30, altura - 40)   # barra superior
    canvas.setFont("Helvetica-Bold", 11)
    canvas.drawString(35, altura - 55, "Lista de Encomendas")
    canvas.drawRightString(largura - 35, altura - 55,
                           "LACOVIANA - Trat. e Lac. Alumínios de Viana, Lda")
    canvas.line(30, altura - 70, largura - 30, altura - 70)   # barra inferior

    canvas.setFont("Helvetica", 8)
    canvas.drawString(35, altura - 85, agora)
    canvas.drawRightString(largura - 35, altura - 85,
                           f"Datas: {filtros.get('data_ini')} a {filtros.get('data_fin')}")
    canvas.drawRightString(largura - 35, altura - 95,
                           f"Clientes: {filtros.get('cliente_ini')} a {filtros.get('cliente_fin')}")
    canvas.drawRightString(largura - 35, altura - 105,
                           f"Tratamentos: {filtros.get('trat_ini')} a {filtros.get('trat_fin')}")
    canvas.line(30, altura - 115, largura - 30, altura - 115)  # separador final do header

def format_num(value):
    try:
        num = float(value)
        if num.is_integer():
            return str(int(num))
        return f"{num:.3f}".rstrip('0').rstrip('.')
    except (TypeError, ValueError):
        return str(value or "")

@app.route("/imprimir", methods=["POST"])
def imprimir():
    filtros = request.form.to_dict()
    resultado = executar_sps(filtros)  # lógica dos SPs já implementada

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            leftMargin=30, rightMargin=30,
                            topMargin=130, bottomMargin=50)

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="BodySmallBold", fontSize=9, leading=11, fontName="Helvetica-Bold"))
    styles.add(ParagraphStyle(name="BodySmall", fontSize=8, leading=10))

    elements = []

    for cliente in resultado:
        cliente_nome = cliente["cliente"].get("cliente", "")
        local = cliente["cliente"].get("local", "")

        for enc in cliente["encomendas"]:
            d = enc["dados"]
            linhas = enc["linhas"]

            flowables = []

            # Cliente + Local
            cliente_table = Table([[
                Paragraph(f"<b>{cliente_nome}</b>", styles["BodySmallBold"]),
                Paragraph(f"<b>{str(local)}</b>", styles["BodySmallBold"])
            ]], colWidths=[300, 200])
            cliente_table.setStyle(TableStyle([
                ('ALIGN', (1,0), (1,0), 'RIGHT'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ]))
            flowables += [cliente_table, Spacer(1,4)]

            # Encomenda
            data_encom = [["Encomenda", "Requisição", "Acabamento", "Micragem", "Conf"], [
                str(d.get("obrano","") or ""),
                str(d.get("obranome","") or ""),
                str(d.get("obs","") or ""),
                str(d.get("micro","") or ""),
                str(d.get("s_n","") or "")
            ]]
            encom_table = Table(data_encom, colWidths=[70, 120, 150, 70, 50], repeatRows=1)
            encom_table.setStyle(TableStyle([
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('FONTSIZE', (0,0), (-1,0), 9),
                ('ALIGN', (0,0), (-1,0), 'CENTER'),
                ('FONTSIZE', (0,1), (-1,-1), 8),
                ('ALIGN', (0,1), (-1,-1), 'LEFT'),
            ]))
            flowables += [encom_table, Spacer(1,4)]

            # Linhas
            if linhas:
                data = [["Artigo", "Descrição", "Qtd", "Medida", "Metros", "Área"]]
                for l in linhas:
                    data.append([
                        str(l.get("ref","") or ""),
                        str(l.get("design","") or ""),
                        format_num(l.get("qtt")),
                        format_num(l.get("u_medida1","")),
                        format_num(l.get("u_mts")),
                        format_num(l.get("u_mts2"))
                    ])
                linhas_table = Table(data, colWidths=[60,200,40,60,60,60], rowHeights=14, repeatRows=1)
                linhas_table.setStyle(TableStyle([
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0,0), (-1,0), 8),
                    ('ALIGN', (0,0), (-1,0), 'CENTER'),
                    ('FONTSIZE', (0,1), (-1,-1), 8),
                    ('ALIGN', (2,1), (-1,-1), 'RIGHT'),
                ]))
                flowables += [linhas_table, Spacer(1,6)]

            # Totais
            totais_table = Table([[f"Total Qtd: {format_num(d.get('qtt',0))}    Total m2: {format_num(d.get('m2',0))}"]],
                                 colWidths=[530])
            totais_table.setStyle(TableStyle([
                ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
                ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
                ('FONTSIZE', (0,0), (-1,-1), 9),
            ]))
            flowables += [totais_table, Spacer(1,6), HRFlowable(width="100%", thickness=0.5, color=colors.black), Spacer(1,8)]

            elements.append(KeepTogether(flowables))

    doc.build(elements,
              onFirstPage=lambda c, d: cabecalho(c, d, filtros),
              onLaterPages=lambda c, d: cabecalho(c, d, filtros),
              canvasmaker=NumberedCanvas)

    pdf = buffer.getvalue()
    buffer.close()

    response = make_response(pdf)
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = 'inline; filename=encomendas.pdf'
    return response



@app.route("/", methods=["GET", "POST"])
def index():
    from datetime import date
    import pyodbc
    import pandas as pd
  
    app_version = ler_versao_local()  # pega versão do version.txt

    # Data inicial default: 1º de janeiro do ano atual
    hoje = date.today()
    data_ini_default = f"{hoje.year}-01-01"
    data_fin_default = hoje.isoformat()       # data de hoje, formato YYYY-MM-DD

    # Conexão a bd
    #conn = pyodbc.connect(conn_str)
    # Ler configurações da BD
    dsn, db, user, password = ler_config()
    if not all([dsn, db, user, password]):
        return redirect(url_for("configuracoes"))
    
    conn = criar_conexao()
    cursor = conn.cursor()

    # Cliente ini/fim defaults
    cursor.execute("SELECT TOP 1 nome FROM cl (NOLOCK) ORDER BY nome ASC")
    row = cursor.fetchone()
    cliente_ini_default = row[0] if row is not None else ""
    
    cursor.execute("SELECT TOP 1 nome FROM cl (NOLOCK) ORDER BY nome DESC")
    row = cursor.fetchone()
    cliente_fin_default = row[0] if row is not None else ""

    # Tratamento ini/fim defaults
    cursor.execute("SELECT TOP 1 tratamento FROM u_tratamentos (NOLOCK) ORDER BY tratamento ASC")
    row = cursor.fetchone()
    trat_ini_default = row[0] if row is not None else ""

    cursor.execute("SELECT TOP 1 tratamento FROM u_tratamentos (NOLOCK) ORDER BY tratamento DESC")
    row = cursor.fetchone()
    trat_fin_default = row[0] if row is not None else ""

    cursor.close()
    conn.close()

    filtros = {
        "data_ini": data_ini_default,
        "data_fin": data_fin_default,
        "cliente_ini":  cliente_ini_default,
        "cliente_fin":  cliente_fin_default,
        "trat_ini": trat_ini_default,
        "trat_fin": trat_fin_default,
        "requisicao": "",
        "enc_ini": 0,
        "enc_fin": 99999999,
        "tipo": "",
        "subtipo": "",
        "gama_cor": "",
        "linha": "",
        "ordem": ""  # default
    }

    if request.method == "POST":
        # Atualiza filtros com os valores do formulário
        for key in filtros.keys():
            filtros[key] = request.form.get(key)  or filtros[key]

    # Inicializa DataFrame e mensagem de erro
    df = pd.DataFrame()
    error_msg = None

    # Chama o SP apenas se for POST
    if request.method == "POST":
        try:
            df, error_msg = get_encomendas(filtros)
        except Exception as e:
            df = pd.DataFrame()
            error_msg = f"Ocorreu um erro inesperado: {e}"

    # Preparar colunas e linhas para a tabela
    columns = list(df.columns) if not df.empty else []
    rows = df.values.tolist() if not df.empty else []

    # Renderiza o template, passando também a mensagem de erro
    return render_template(
        "index.html",
        columns=columns,
        rows=rows,
        filtros=filtros,
        error_msg=error_msg,
        app_version=app_version
    )

import threading
import webview

if __name__ == "__main__":
    port = 5000

    def start_flask():
        app.run(debug=False, port=port, use_reloader=False)

    # Corre Flask em segundo plano
    threading.Thread(target=start_flask, daemon=True).start()

    # Abre numa janela nativa (sem precisar do Chrome/Edge)
    webview.create_window("Encomendas", f"http://127.0.0.1:{port}")
    webview.start()


