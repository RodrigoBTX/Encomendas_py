from flask import Flask, render_template, request, jsonify,redirect, url_for, session
import pyodbc
import pandas as pd
import configparser
import os
import base64
import webbrowser
from threading import Timer

app = Flask(__name__)
app.secret_key = os.urandom(24)  # Chave segura para sessões

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
# conn_str = "DSN=Lacoviana;UID=admin-btx;PWD=093N3mmb!;DATABASE=PHC_20240310"
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
ADMIN_PASSWORD = "093N3mmb!"

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




@app.route("/", methods=["GET", "POST"])
def index():
    from datetime import date
    import pyodbc
    import pandas as pd
    # Data inicial default: 1º de janeiro do ano atual
    hoje = date.today()
    data_ini_default = f"{hoje.year}-01-01"
    data_fin_default = hoje.isoformat()       # data de hoje, formato YYYY-MM-DD

    # Conexão ao banco
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
        error_msg=error_msg
    )

import webbrowser
from threading import Timer

if __name__ == "__main__":
    port = 5000
    # Abrir navegador automaticamente
    Timer(1, lambda: webbrowser.open(f"http://127.0.0.1:{port}")).start()
    app.run(debug=False, port=port)

