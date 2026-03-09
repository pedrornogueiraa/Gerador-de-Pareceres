"""
Criador de Pareceres Atuariais - Backend
"""

from flask import Flask, jsonify, request, send_file, send_from_directory, session
from flask_cors import CORS
from functools import wraps
import dropbox
from dropbox.exceptions import AuthError, ApiError
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy
import io
import os


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
app.secret_key = os.environ.get("SECRET_KEY", "lumens-secret-2024")
CORS(app, supports_credentials=True)

TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template.docx')


# Autenticacao

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('autenticado'):
            return jsonify({"erro": "Nao autenticado."}), 401
        return f(*args, **kwargs)
    return decorated


@app.route("/api/login", methods=["POST"])
def rota_login():
    data = request.json
    senha = data.get("senha", "")
    senha_correta = os.environ.get("APP_SENHA", "lumens2024")
    if senha == senha_correta:
        session['autenticado'] = True
        return jsonify({"ok": True})
    return jsonify({"erro": "Senha incorreta."}), 401


@app.route("/api/logout", methods=["POST"])
def rota_logout():
    session.clear()
    return jsonify({"ok": True})


@app.route("/api/check", methods=["GET"])
def rota_check():
    return jsonify({"autenticado": bool(session.get('autenticado'))})


# Dropbox

def get_dropbox_client():
    refresh_token = os.environ.get("DROPBOX_REFRESH_TOKEN", "").strip()
    app_key = os.environ.get("DROPBOX_APP_KEY", "").strip()
    app_secret = os.environ.get("DROPBOX_APP_SECRET", "").strip()

    if refresh_token and app_key and app_secret:
        return dropbox.Dropbox(
            oauth2_refresh_token=refresh_token,
            app_key=app_key,
            app_secret=app_secret
        )

    token = os.environ.get("DROPBOX_TOKEN", "").strip()
    if token:
        return dropbox.Dropbox(token)

    raise ValueError("Credenciais do Dropbox nao configuradas.")


def get_pasta():
    return os.environ.get("DROPBOX_PASTA", "/Banco de Teses").strip()


# Rotas

@app.route("/")
def index():
    pasta = os.path.dirname(os.path.abspath(__file__))
    return send_from_directory(pasta, 'criador-pareceres.html')


@app.route("/api/status", methods=["GET"])
def status():
    refresh_token = os.environ.get("DROPBOX_REFRESH_TOKEN", "").strip()
    app_key = os.environ.get("DROPBOX_APP_KEY", "").strip()
    token = os.environ.get("DROPBOX_TOKEN", "").strip()
    dropbox_ok = bool((refresh_token and app_key) or token)
    template_ok = os.path.exists(TEMPLATE_PATH)
    return jsonify({
        "status": "ok",
        "dropbox_configurado": dropbox_ok,
        "pasta": get_pasta(),
        "template_ok": template_ok
    })


@app.route("/api/explorar", methods=["GET"])
def explorar():
    try:
        dbx = get_dropbox_client()
        path = request.args.get("path", "")
        resultado = dbx.files_list_folder(path, recursive=False)
        itens = [{"nome": e.name, "path": e.path_display} for e in resultado.entries]
        return jsonify({"path_consultado": path if path else "/", "itens": itens})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500


@app.route("/api/topicos", methods=["GET"])
@login_required
def listar_topicos():
    try:
        dbx = get_dropbox_client()
        pasta = get_pasta()
        resultado = dbx.files_list_folder(pasta, recursive=True)
        topicos = []
        idx = 1
        for entry in resultado.entries:
            if isinstance(entry, dropbox.files.FileMetadata):
                if entry.name.lower().endswith(".docx"):
                    partes = entry.path_lower.replace(pasta.lower() + "/", "").split("/")
                    categoria = partes[0].title() if len(partes) > 1 else "Geral"
                    nome = entry.name.replace(".docx", "").replace(".DOCX", "")
                    topicos.append({
                        "id": idx,
                        "nome": nome,
                        "categoria": categoria,
                        "path": entry.path_display,
                    })
                    idx += 1
        topicos.sort(key=lambda x: (x["categoria"], x["nome"]))
        return jsonify({"topicos": topicos})
    except AuthError:
        return jsonify({"erro": "Token invalido ou expirado."}), 401
    except ApiError as e:
        return jsonify({"erro": f"Erro na API do Dropbox: {str(e)}"}), 500
    except ValueError as e:
        return jsonify({"erro": str(e)}), 400
    except Exception as e:
        return jsonify({"erro": f"Erro inesperado: {str(e)}"}), 500


@app.route("/api/gerar", methods=["POST"])
@login_required
def gerar_parecer():
    data = request.json
    topicos_selecionados = data.get("topicos", [])
    dados_caso = data.get("dados", {})

    if not topicos_selecionados:
        return jsonify({"erro": "Nenhum topico selecionado."}), 400

    if not os.path.exists(TEMPLATE_PATH):
        return jsonify({"erro": "Arquivo template.docx nao encontrado."}), 500

    try:
        dbx = get_dropbox_client()

        doc = Document(TEMPLATE_PATH)
        _substituir_dados_doc(doc, dados_caso)

        placeholder_elem = _encontrar_placeholder_primeira_impugnacao(doc)
        ponto_insercao = _encontrar_ponto_insercao(doc)

        for i, topico in enumerate(topicos_selecionados):
            docx_topico = _baixar_docx(dbx, topico["path"])

            if i == 0 and placeholder_elem is not None:
                ref_element = placeholder_elem
                for elem in docx_topico.element.body:
                    if elem.tag == qn("w:sectPr"):
                        continue
                    novo = copy.deepcopy(elem)
                    _substituir_dados_xml(novo, dados_caso)
                    ref_element.addprevious(novo)
                ref_element.getparent().remove(ref_element)
            else:
                ref_element = ponto_insercao
                sep = _criar_paragrafo_vazio()
                ref_element.addprevious(sep)
                for elem in docx_topico.element.body:
                    if elem.tag == qn("w:sectPr"):
                        continue
                    novo = copy.deepcopy(elem)
                    _substituir_dados_xml(novo, dados_caso)
                    ref_element.addprevious(novo)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        processo = dados_caso.get("processo", "sp").strip().replace("/", "-").replace(".", "-").replace(" ", "_")
        nome_arquivo = f"{processo}_Parecer Tecnico.rev001.docx"

        return send_file(
            buffer,
            as_attachment=True,
            download_name=nome_arquivo,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except AuthError:
        return jsonify({"erro": "Token invalido ou expirado."}), 401
    except Exception as e:
        import traceback
        return jsonify({"erro": f"Erro ao gerar parecer: {str(e)}", "detalhe": traceback.format_exc()}), 500


# Auxiliares

def _encontrar_placeholder_primeira_impugnacao(doc):
    for elem in doc.element.body:
        if elem.tag == qn("w:p"):
            texto = "".join(t.text or "" for t in elem.iter(qn("w:t")))
            if "inserir primeira impugna" in texto.lower():
                return elem
    return None


def _encontrar_ponto_insercao(doc):
    body = doc.element.body
    nivel1_encontrados = 0
    for elem in body:
        if elem.tag == qn("w:p"):
            pPr = elem.find(qn("w:pPr"))
            if pPr is not None:
                pStyle = pPr.find(qn("w:pStyle"))
                if pStyle is not None:
                    if pStyle.get(qn("w:val"), "") == "cTTULONVEL1":
                        nivel1_encontrados += 1
                        if nivel1_encontrados == 2:
                            return elem
    sect_pr = body.find(qn("w:sectPr"))
    if sect_pr is not None:
        return sect_pr
    return None


def _criar_paragrafo_vazio():
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    pStyle = OxmlElement("w:pStyle")
    pStyle.set(qn("w:val"), "fTextodocorpo")
    pPr.append(pStyle)
    p.append(pPr)
    return p


def _baixar_docx(dbx, path):
    _, response = dbx.files_download(path)
    return Document(io.BytesIO(response.content))


def _substituir_dados_doc(doc, dados_caso):
    for p in doc.paragraphs:
        _consolidar_e_substituir(p, dados_caso)


def _consolidar_e_substituir(p, dados_caso):
    texto_completo = p.text
    mapeamento = _build_mapeamento(dados_caso)
    if not any(ph in texto_completo for ph in mapeamento):
        return
    novo_texto = texto_completo
    for ph, val in mapeamento.items():
        novo_texto = novo_texto.replace(ph, val)
    runs = p.runs
    if not runs:
        return
    runs[0].text = novo_texto
    for r in runs[1:]:
        r.text = ""


def _substituir_dados_xml(elemento, dados_caso):
    mapeamento = _build_mapeamento(dados_caso)
    for no_texto in elemento.iter(qn("w:t")):
        if no_texto.text:
            for ph, val in mapeamento.items():
                if ph in no_texto.text:
                    no_texto.text = no_texto.text.replace(ph, val)


def _build_mapeamento(dados_caso):
    mapa = {
        "[demanda]":           dados_caso.get("demanda", ""),
        "[nº do processo]":    dados_caso.get("processo", ""),
        "[Autor/Reclamante]":  dados_caso.get("participante", ""),
        "[Vara/Juízo]":        dados_caso.get("vara", ""),
        "[data de entrega]":   dados_caso.get("entrega", ""),
    }
    return {k: v for k, v in mapa.items() if v}


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
