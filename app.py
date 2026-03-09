"""
Criador de Pareceres Atuariais - Backend
"""

from flask import Flask, jsonify, request, send_file, send_from_directory
from flask_cors import CORS
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
CORS(app)

# Caminho do template base (deve estar na mesma pasta que app.py)
TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template.docx')


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


@app.route("/")
def index():
    pasta = os.path.dirname(os.path.abspath(__file__))
    return send_from_directory(pasta, 'criador-pareceres.html')


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
def gerar_parecer():
    data = request.json
    topicos_selecionados = data.get("topicos", [])
    dados_caso = data.get("dados", {})

    if not topicos_selecionados:
        return jsonify({"erro": "Nenhum topico selecionado."}), 400

    if not os.path.exists(TEMPLATE_PATH):
        return jsonify({"erro": "Arquivo template.docx nao encontrado na pasta do servidor."}), 500

    try:
        dbx = get_dropbox_client()

        # 1. Carrega o template como base
        doc = Document(TEMPLATE_PATH)

        # 2. Substitui dados da capa
        _substituir_dados_doc(doc, dados_caso)

        # 3. Encontra o ponto de inserção dos tópicos
        #    Estratégia: insere após o último parágrafo com estilo dTTULONVEL2 que vier
        #    depois do primeiro cTTULONVEL1, e antes do próximo cTTULONVEL1 ("Do cálculo")
        ponto_insercao = _encontrar_ponto_insercao(doc)

        # 4. Baixa e insere cada tópico na posição correta
        body = doc.element.body
        ref_element = ponto_insercao  # insere antes deste elemento

        for i, topico in enumerate(topicos_selecionados):
            docx_topico = _baixar_docx(dbx, topico["path"])

            # Opcional: adiciona parágrafo separador entre tópicos
            if i > 0:
                sep = _criar_paragrafo_vazio()
                ref_element.addprevious(sep)

            # Copia todos os elementos do tópico (exceto sectPr)
            for elem in docx_topico.element.body:
                if elem.tag == qn("w:sectPr"):
                    continue
                novo = copy.deepcopy(elem)
                _substituir_dados_xml(novo, dados_caso)
                ref_element.addprevious(novo)

        # 5. Salva e retorna
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        participante = dados_caso.get("participante", "Participante").split()[0]
        processo = dados_caso.get("processo", "").replace("/", "-").replace(".", "")[:15]
        nome_arquivo = f"Parecer_{participante}_{processo}.docx".replace(" ", "_")

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


def _encontrar_ponto_insercao(doc):
    """
    Encontra o elemento antes do qual os tópicos devem ser inseridos.
    Lógica: após o primeiro cTTULONVEL1, procura o segundo cTTULONVEL1
    (que normalmente é 'Do cálculo da PREVI' ou similar).
    Insere ANTES do segundo cTTULONVEL1.
    Se não encontrar, insere antes do sectPr final.
    """
    body = doc.element.body
    nivel1_encontrados = 0

    for elem in body:
        # Verifica se é parágrafo com estilo cTTULONVEL1
        if elem.tag == qn("w:p"):
            pPr = elem.find(qn("w:pPr"))
            if pPr is not None:
                pStyle = pPr.find(qn("w:pStyle"))
                if pStyle is not None:
                    style_val = pStyle.get(qn("w:val"), "")
                    if style_val == "cTTULONVEL1":
                        nivel1_encontrados += 1
                        if nivel1_encontrados == 2:
                            return elem  # insere ANTES deste

    # Fallback: insere antes do sectPr
    sect_pr = body.find(qn("w:sectPr"))
    if sect_pr is not None:
        return sect_pr

    return None


def _criar_paragrafo_vazio():
    """Cria parágrafo vazio com estilo fTextodocorpo."""
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
    """Substitui placeholders em todos os parágrafos do documento."""
    for p in doc.paragraphs:
        for run in p.runs:
            _substituir_texto_run(run, dados_caso)


def _substituir_dados_xml(elemento, dados_caso):
    """Substitui placeholders nos nós w:t de um elemento XML."""
    mapeamento = _build_mapeamento(dados_caso)
    for no_texto in elemento.iter(qn("w:t")):
        if no_texto.text:
            for ph, val in mapeamento.items():
                if ph in no_texto.text and val:
                    no_texto.text = no_texto.text.replace(ph, val)


def _substituir_texto_run(run, dados_caso):
    mapeamento = _build_mapeamento(dados_caso)
    for ph, val in mapeamento.items():
        if ph in run.text and val:
            run.text = run.text.replace(ph, val)


def _build_mapeamento(dados_caso):
    return {
        "[demanda]":           dados_caso.get("demanda", ""),
        "[nº do processo]":    dados_caso.get("processo", ""),
        "[Autor/Reclamante]":  dados_caso.get("participante", ""),
        "[Vara/Juízo]":        dados_caso.get("vara", ""),
        "[data de entrega]":   dados_caso.get("entrega", ""),
    }


@app.route("/api/status", methods=["GET"])
def status():
    token = os.environ.get("DROPBOX_TOKEN", "").strip()
    template_ok = os.path.exists(TEMPLATE_PATH)
    return jsonify({
        "status": "ok",
        "dropbox_configurado": bool(token),
        "pasta": get_pasta(),
        "template_ok": template_ok
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
