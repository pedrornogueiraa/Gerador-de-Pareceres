"""
Gerador de Pareceres Atuariais — Backend
Lumens Atuarial | Núcleo Judicial
"""

from flask import Flask, jsonify, request, send_file, send_from_directory, session
from flask_cors import CORS
import dropbox
from dropbox.exceptions import AuthError, ApiError
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy
import io
import os
import re
import functools

# ── ENV ───────────────────────────────────────────────────────────────────────

def carregar_env():
    env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.env')
    if os.path.exists(env_path):
        with open(env_path, 'r', encoding='utf-8') as f:
            for linha in f:
                linha = linha.strip()
                if linha and '=' in linha and not linha.startswith('#'):
                    chave, valor = linha.split('=', 1)
                    os.environ[chave.strip()] = valor.strip()

carregar_env()

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "lumens-judicial-chave-2026")
CORS(app, supports_credentials=True, origins=[
    "https://criador-de-pareceres.onrender.com",
    "https://gerador-de-pareceres.onrender.com",
    "http://localhost:5000"
])

# ── HELPERS ───────────────────────────────────────────────────────────────────

def get_dropbox_client():
    refresh_token = os.environ.get("DROPBOX_REFRESH_TOKEN", "").strip()
    app_key       = os.environ.get("DROPBOX_APP_KEY", "").strip()
    app_secret    = os.environ.get("DROPBOX_APP_SECRET", "").strip()
    if refresh_token and app_key and app_secret:
        return dropbox.Dropbox(
            oauth2_refresh_token=refresh_token,
            app_key=app_key,
            app_secret=app_secret
        )
    token = os.environ.get("DROPBOX_TOKEN", "").strip()
    if token:
        return dropbox.Dropbox(token)
    raise ValueError("Credenciais do Dropbox não configuradas.")

def get_pasta_banco():
    return os.environ.get("DROPBOX_PASTA", "/Banco de Teses").strip()

def get_pasta_estrutura():
    return os.environ.get("DROPBOX_PASTA_ESTRUTURA", "/Estrutura").strip()

def login_required(f):
    @functools.wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("autenticado"):
            return jsonify({"erro": "Não autenticado"}), 401
        return f(*args, **kwargs)
    return decorated

def _listar_pasta_dropbox(dbx, pasta):
    """Lista todos os .docx de uma pasta recursivamente com paginação completa."""
    resultado = dbx.files_list_folder(pasta, recursive=True)
    entradas  = list(resultado.entries)
    while resultado.has_more:
        resultado = dbx.files_list_folder_continue(resultado.cursor)
        entradas.extend(resultado.entries)
    return entradas

def _extrair_topicos(entradas, pasta_raiz, ordem_prioridade=None):
    """
    Converte entradas do Dropbox em lista de tópicos.
    ordem_prioridade: lista de nomes de subpasta (lowercase) que vêm primeiro.
    """
    topicos = []
    idx     = 1
    prefixo = pasta_raiz.rstrip("/") + "/"

    for entry in entradas:
        if not isinstance(entry, dropbox.files.FileMetadata):
            continue
        if not entry.name.lower().endswith(".docx"):
            continue

        pd = entry.path_display
        rel = pd[len(prefixo):] if pd.lower().startswith(prefixo.lower()) else pd
        partes = rel.split("/")

        if len(partes) > 1:
            categoria_display = partes[0]
            categoria_lower   = partes[0].lower()
        else:
            categoria_display = "Geral"
            categoria_lower   = "geral"

        nome = entry.name
        for ext in (".docx", ".DOCX", ".Docx"):
            nome = nome.replace(ext, "")

        eh_ultima_pagina = nome.lower() in ("anexo", "apêndice", "apendice")

        topicos.append({
            "id":            idx,
            "nome":          nome,
            "categoria":     categoria_display,
            "categoria_lower": categoria_lower,
            "path":          entry.path_display,
            "ultima_pagina": eh_ultima_pagina,
        })
        idx += 1

    def sort_key(x):
        cat = x["categoria_lower"]
        if ordem_prioridade and cat in ordem_prioridade:
            return (0, ordem_prioridade.index(cat), x["nome"].lower())
        return (1, 999, cat + x["nome"].lower())

    topicos.sort(key=sort_key)
    return topicos

# ── AUTENTICAÇÃO ──────────────────────────────────────────────────────────────

@app.route("/api/check", methods=["GET"])
def check():
    return jsonify({"autenticado": bool(session.get("autenticado"))})

@app.route("/api/login", methods=["POST"])
def rota_login():
    data  = request.get_json() or {}
    senha = data.get("senha", "")
    senha_correta = os.environ.get("APP_SENHA", "JudicialLumens01")
    if senha == senha_correta:
        session["autenticado"] = True
        session.permanent = False
        return jsonify({"ok": True})
    return jsonify({"erro": "Senha incorreta"}), 401

@app.route("/api/logout", methods=["POST"])
def rota_logout():
    session.clear()
    return jsonify({"ok": True})

# ── FRONTEND ──────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    pasta = os.path.dirname(os.path.abspath(__file__))
    return send_from_directory(pasta, 'criador-pareceres.html')

# ── STATUS ────────────────────────────────────────────────────────────────────

@app.route("/api/status", methods=["GET"])
def status():
    template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template.docx')
    dropbox_ok = bool(
        os.environ.get("DROPBOX_REFRESH_TOKEN") or os.environ.get("DROPBOX_TOKEN")
    )
    return jsonify({
        "status":            "ok",
        "dropbox_configurado": dropbox_ok,
        "template_ok":       os.path.exists(template_path),
        "pasta":             get_pasta_banco()
    })

# ── DIAGNÓSTICO ───────────────────────────────────────────────────────────────

@app.route("/api/explorar", methods=["GET"])
def explorar():
    try:
        dbx  = get_dropbox_client()
        path = request.args.get("path", "")
        resultado = dbx.files_list_folder(path, recursive=False)
        itens = [{"nome": e.name, "path": e.path_display} for e in resultado.entries]
        return jsonify({"path_consultado": path if path else "/", "itens": itens})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500

# ── ESTRUTURA ─────────────────────────────────────────────────────────────────

@app.route("/api/estrutura", methods=["GET"])
@login_required
def listar_estrutura():
    """Retorna arquivos da pasta /Estrutura — subpastas em ordem alfabética."""
    try:
        dbx     = get_dropbox_client()
        pasta   = get_pasta_estrutura()
        entradas = _listar_pasta_dropbox(dbx, pasta)
        topicos  = _extrair_topicos(entradas, pasta, ordem_prioridade=None)
        return jsonify({"topicos": topicos})
    except AuthError:
        return jsonify({"erro": "Token do Dropbox inválido ou expirado."}), 401
    except ApiError as e:
        return jsonify({"erro": f"Erro na API do Dropbox: {str(e)}"}), 500
    except Exception as e:
        return jsonify({"erro": f"Erro inesperado: {str(e)}"}), 500

# ── BANCO DE TESES ────────────────────────────────────────────────────────────

@app.route("/api/topicos", methods=["GET"])
@login_required
def listar_topicos():
    """
    Retorna arquivos do Banco de Teses do cliente selecionado.
    Parâmetro: ?cliente=PREVI  ou  ?cliente=ELOS
    Estrutura: /Banco de Teses/{CLIENTE}/Banco de Teses {CLIENTE}/subpastas/
    """
    try:
        dbx     = get_dropbox_client()
        cliente = request.args.get("cliente", "PREVI").strip().upper()

        # Validar cliente
        if cliente not in ("PREVI", "ELOS"):
            return jsonify({"erro": "Cliente inválido. Use PREVI ou ELOS."}), 400

        pasta_raiz    = get_pasta_banco()
        # Estrutura real: /Banco de Teses/Banco de Teses PREVI/
        pasta_cliente = f"{pasta_raiz}/Banco de Teses {cliente}"

        entradas = _listar_pasta_dropbox(dbx, pasta_cliente)
        topicos  = _extrair_topicos(entradas, pasta_cliente, ordem_prioridade=None)

        return jsonify({"topicos": topicos, "cliente": cliente})

    except AuthError:
        return jsonify({"erro": "Token do Dropbox inválido ou expirado."}), 401
    except ApiError as e:
        return jsonify({"erro": f"Erro na API do Dropbox: {str(e)}"}), 500
    except ValueError as e:
        return jsonify({"erro": str(e)}), 400
    except Exception as e:
        return jsonify({"erro": f"Erro inesperado: {str(e)}"}), 500

# ── GERAÇÃO DO PARECER ────────────────────────────────────────────────────────

@app.route("/api/gerar", methods=["POST"])
@login_required
def gerar_parecer():
    data                 = request.get_json() or {}
    topicos_selecionados = data.get("topicos", [])
    dados_caso           = data.get("dados", {})

    if not topicos_selecionados:
        return jsonify({"erro": "Nenhum tópico selecionado."}), 400

    try:
        dbx           = get_dropbox_client()
        template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template.docx')

        if not os.path.exists(template_path):
            return jsonify({"erro": "template.docx não encontrado no servidor."}), 500

        doc_final = Document(template_path)
        _substituir_dados(doc_final.element.body, dados_caso)

        normais    = [t for t in topicos_selecionados if not t.get("ultima_pagina")]
        ultima_pag = [t for t in topicos_selecionados if t.get("ultima_pagina")]

        ponto = _encontrar_ponto_insercao(doc_final)

        for topico in normais:
            docx_topico  = _baixar_docx(dbx, topico["path"])
            _copiar_estilos(docx_topico, doc_final)
            eh_principal = topico.get("topico_principal", False)
            _inserir_topico(docx_topico, doc_final, dados_caso, ponto, eh_principal)
            ponto = None  # próximos entram em sequência antes do encerramento

        for topico in ultima_pag:
            docx_topico = _baixar_docx(dbx, topico["path"])
            _copiar_estilos(docx_topico, doc_final)
            _inserir_ultima_pagina(docx_topico, doc_final, dados_caso)

        buffer = io.BytesIO()
        doc_final.save(buffer)
        buffer.seek(0)

        processo      = dados_caso.get("processo", "sp").strip()
        processo_limpo = re.sub(r'[\\/*?:"<>|]', '-', processo)
        nome_arquivo  = f"{processo_limpo}_Parecer Técnico.rev001.docx"

        return send_file(
            buffer,
            as_attachment=True,
            download_name=nome_arquivo,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except AuthError:
        return jsonify({"erro": "Token do Dropbox inválido ou expirado."}), 401
    except Exception as e:
        import traceback
        return jsonify({"erro": f"Erro ao gerar parecer: {str(e)}", "detalhe": traceback.format_exc()}), 500

# ── FUNÇÕES AUXILIARES ────────────────────────────────────────────────────────

def _baixar_docx(dbx, path):
    _, response = dbx.files_download(path)
    return Document(io.BytesIO(response.content))


def _copiar_estilos(doc_origem, doc_destino):
    estilos_origem  = doc_origem.element.find(qn("w:styles"))
    estilos_destino = doc_destino.element.find(qn("w:styles"))
    if estilos_origem is None or estilos_destino is None:
        return
    ids_existentes = {
        e.get(qn("w:styleId"))
        for e in estilos_destino.findall(qn("w:style"))
        if e.get(qn("w:styleId"))
    }
    for estilo in estilos_origem.findall(qn("w:style")):
        sid = estilo.get(qn("w:styleId"))
        if sid and sid not in ids_existentes:
            estilos_destino.append(copy.deepcopy(estilo))


def _encontrar_ponto_insercao(doc):
    body     = doc.element.body
    children = list(body)

    # Prioridade 1: placeholder explícito
    for elem in children:
        for t in elem.iter(qn("w:t")):
            if t.text and "[inserir primeira impugna" in t.text:
                return elem

    # Prioridade 2: elemento imediatamente antes de "É o Parecer Técnico."
    for i, elem in enumerate(children):
        for t in elem.iter(qn("w:t")):
            if t.text and "É o Parecer Técnico" in t.text:
                return children[i - 1] if i > 0 else elem

    # Fallback: antes do sectPr
    sect_pr = body.find(qn("w:sectPr"))
    if sect_pr is not None:
        idx = children.index(sect_pr)
        return children[idx - 1] if idx > 0 else None
    return None


def _encontrar_inicio_encerramento(doc):
    body     = doc.element.body
    children = list(body)
    for i, elem in enumerate(children):
        for t in elem.iter(qn("w:t")):
            if t.text and "É o Parecer Técnico" in t.text:
                if i > 0:
                    prev_texts = [t.text for t in children[i-1].iter(qn("w:t")) if t.text]
                    if not prev_texts:
                        return i - 1
                return i
    return None


def _encontrar_pos_encerramento(doc):
    body     = doc.element.body
    children = list(body)
    em_encerramento = False
    for i, elem in enumerate(children):
        for t in elem.iter(qn("w:t")):
            if t.text and "É o Parecer Técnico" in t.text:
                em_encerramento = True
        if em_encerramento:
            tem_page_break = any(
                br.get(qn("w:type")) == "page"
                for br in elem.iter(qn("w:br"))
            )
            if tem_page_break:
                return i + 1
    return None


def _inserir_topico(doc_topico, doc_final, dados_caso, ponto_ref, eh_principal):
    body_origem  = doc_topico.element.body
    body_destino = doc_final.element.body
    elementos    = [e for e in body_origem if e.tag != qn("w:sectPr")]
    if not elementos:
        return

    if ponto_ref is not None and ponto_ref in list(body_destino):
        pos = list(body_destino).index(ponto_ref)
        placeholder_encontrado = any(
            "[inserir primeira impugna" in (t.text or "")
            for t in ponto_ref.iter(qn("w:t"))
        )
        if placeholder_encontrado:
            insert_pos = pos
            body_destino.remove(ponto_ref)
        else:
            insert_pos = pos + 1
    else:
        idx_enc = _encontrar_inicio_encerramento(doc_final)
        if idx_enc is not None:
            insert_pos = idx_enc
        else:
            sect_pr    = body_destino.find(qn("w:sectPr"))
            insert_pos = list(body_destino).index(sect_pr) if sect_pr is not None else len(list(body_destino))

    for i, elem in enumerate(elementos):
        novo = copy.deepcopy(elem)
        _substituir_dados(novo, dados_caso)
        if i == 0:
            _ajustar_estilo_titulo(novo, eh_principal)
        body_destino.insert(insert_pos + i, novo)


def _ajustar_estilo_titulo(elem, eh_principal):
    pPr = elem.find(qn("w:pPr"))
    if pPr is None:
        return
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is None:
        return
    estilo_atual = pStyle.get(qn("w:val"), "")
    estilos_titulo = {"cTTULONVEL1", "dTTULONVEL2", "Heading1", "Heading2"}
    if estilo_atual in estilos_titulo or "TTULO" in estilo_atual.upper():
        pStyle.set(qn("w:val"), "cTTULONVEL1" if eh_principal else "dTTULONVEL2")


def _inserir_ultima_pagina(doc_topico, doc_final, dados_caso):
    body_origem  = doc_topico.element.body
    body_destino = doc_final.element.body
    sect_pr      = body_destino.find(qn("w:sectPr"))

    pos_enc = _encontrar_pos_encerramento(doc_final)
    if pos_enc is None:
        children = list(body_destino)
        pos_enc  = children.index(sect_pr) if sect_pr is not None else len(children)

    offset = 0
    for elem in body_origem:
        if elem.tag == qn("w:sectPr"):
            continue
        novo = copy.deepcopy(elem)
        _substituir_dados(novo, dados_caso)
        body_destino.insert(pos_enc + offset, novo)
        offset += 1


def _substituir_dados(elemento, dados_caso):
    mapeamento = {
        "[demanda]":          dados_caso.get("demanda", ""),
        "[nº do processo]":   dados_caso.get("processo", ""),
        "[dossiê]":           dados_caso.get("dossie", ""),
        "[Autor/Reclamante]": dados_caso.get("participante", ""),
        "[Vara/Juízo]":       dados_caso.get("vara", ""),
        "[data de entrega]":  dados_caso.get("entrega", ""),
        "{{participante}}":   dados_caso.get("participante", ""),
        "{{processo}}":       dados_caso.get("processo", ""),
        "{{vara}}":           dados_caso.get("vara", ""),
        "{{demanda}}":        dados_caso.get("demanda", ""),
        "{{entrega}}":        dados_caso.get("entrega", ""),
    }
    for no_texto in elemento.iter(qn("w:t")):
        if no_texto.text:
            for placeholder, valor in mapeamento.items():
                if placeholder in no_texto.text and valor:
                    no_texto.text = no_texto.text.replace(placeholder, valor)


# ── INICIALIZAÇÃO ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
