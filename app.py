#!/usr/bin/env python3
"""
Web app: login -> faturas na tela + "Exportar PDF" (mesmo PDF do relatorio_faturas_pdf.py)

Local:   APP_PASSWORD=xxx python app.py        -> http://localhost:5000
Render:  gunicorn app:app   (ver render.yaml; precisa ser 1 worker: os jobs ficam na memória)

Variáveis: APP_PASSWORD + as TOTVS_* do .env (no Render, em Environment).
SECRET_KEY é opcional: sem ela o login cai sempre que o servidor reinicia.

Fluxo (igual ao relatorio-credev): "Consultar" abre /?venc_ini=...&venc_fim=..., que faz
POST /carregar; a busca roda numa thread e o navegador consulta /progresso/<job> a cada
segundo (barra de %). Pronto, pega a tabela em /tabela/<job>.

O PDF só é montado quando pedido (POST /pdf/<job>/gerar), com os mesmos dados da tela, e depois
sai em /pdf/<job>: em períodos grandes ele leva minutos, e a tabela não pode esperar por isso.
"""

from __future__ import annotations

import hmac
import io
import math
import os
import secrets
import sys
import threading
import time
import traceback
from datetime import date, datetime

from flask import Flask, abort, jsonify, redirect, render_template, request, send_file, session, url_for

# os scripts imprimem "✔"/"▶": não deixar um console sem UTF-8 derrubar a geração
for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(errors="replace")
    except AttributeError:
        pass

# os scripts leem sys.argv ao serem importados; sob o gunicorn os argumentos são dele
_argv, sys.argv = sys.argv, sys.argv[:1]
import relatorio_faturas_pdf as PDF  # noqa: E402
sys.argv = _argv

S = PDF.S
R = PDF.RF.R
APP_PASSWORD = os.getenv("APP_PASSWORD", "")
TRAVA = threading.Lock()  # consulta e PDF mexem em variáveis globais dos scripts: um por vez
JOBS: dict[str, dict] = {}
JOB_TTL = 30 * 60         # consultas ficam disponíveis por 30 min

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY") or os.urandom(32)
app.add_template_filter(PDF.moeda, "moeda")
app.add_template_filter(PDF.dt, "data")
app.add_template_filter(PDF.milhar, "milhar")

# ----------------------------------------------------------------------------
# Progresso: etapas com faixas de % (títulos 2-45, notas 45-75, clientes 75-92, relatório 92-100)
# ----------------------------------------------------------------------------

ATUAL: dict = {}  # job em andamento (só um por vez, por causa da TRAVA)


def _progresso(rotulo: str, pagina: int, total: int) -> None:
    j = ATUAL
    frac = min(pagina / max(total, 1), 1)
    if rotulo == "títulos":
        j["pct"], j["etapa"] = 2 + 43 * frac, f"Buscando títulos (página {pagina}/{total})"
    elif rotulo.startswith("notas"):  # "notas" ou "notas 2/3" (consulta em janelas de 6 meses)
        parte, partes = (int(x) for x in rotulo.split()[1].split("/")) if " " in rotulo else (1, 1)
        frac = (parte - 1 + frac) / partes
        sufixo = f", parte {parte}/{partes}" if partes > 1 else ""
        j["pct"], j["etapa"] = 45 + 30 * frac, f"Buscando notas (página {pagina}/{total}{sufixo})"
    elif rotulo.startswith("pessoas"):
        if pagina >= total:
            j["pess_feitas"] = j.get("pess_feitas", 0) + 1
        feitas, tot = j.get("pess_feitas", 0), max(j.get("pess_total", 1), 1)
        j["pct"], j["etapa"] = 75 + 17 * min(feitas / tot, 1), "Buscando cadastro dos clientes"


_buscar_pessoas_orig = R.buscar_pessoas


def _buscar_pessoas(api, codigos):
    ATUAL["pess_total"] = 2 * math.ceil(len(codigos) / 50)  # lotes de 50, PJ + PF
    ATUAL["pct"], ATUAL["etapa"] = 75, "Buscando cadastro dos clientes"
    resultado = _buscar_pessoas_orig(api, codigos)
    ATUAL["pct"], ATUAL["etapa"] = 92, "Montando o relatório"
    return resultado


R.buscar_pessoas = _buscar_pessoas
R.PROGRESSO = _progresso


def _rodar(job: dict, d_ini: date, d_fim: date) -> None:
    global ATUAL
    try:
        ATUAL = job
        job["etapa"] = "Conectando ao TOTVS"
        S.VENC_INI, S.VENC_FIM = d_ini.isoformat(), d_fim.isoformat()
        job["r"] = PDF.obter_relatorio()
        job["pct"], job["etapa"], job["pronto"] = 100, "Pronto", True
    except SystemExit as e:  # os scripts usam sys.exit("mensagem") para erros esperados
        job["erro"] = str(e)
    except Exception as e:  # noqa: BLE001
        traceback.print_exc()
        job["erro"] = f"Erro ao montar o relatório: {e}"
    finally:
        ATUAL = {}
        TRAVA.release()


def _rodar_pdf(job: dict) -> None:
    try:
        S.VENC_INI, S.VENC_FIM = job["ini"].isoformat(), job["fim"].isoformat()  # cabeçalho e nome do arquivo
        agora = datetime.now()
        buf = io.BytesIO()
        PDF.escrever_pdf(buf, job["r"], agora, lambda pagina: job.update(pdf_pag=pagina))
        job["nome"], job["pdf"] = PDF.nome_arquivo(agora), buf.getvalue()
    except Exception as e:  # noqa: BLE001
        traceback.print_exc()
        job["pdf_erro"] = f"Erro ao gerar o PDF: {e}"
    finally:
        job["pdf_gerando"] = False
        TRAVA.release()


limite_6_meses = PDF.RF.limite_6_meses


def _texto(msg: str, status: int):
    return msg, status, {"Content-Type": "text/plain; charset=utf-8"}


# ----------------------------------------------------------------------------
# Rotas
# ----------------------------------------------------------------------------

@app.before_request
def exigir_senha():
    if request.endpoint != "login" and not session.get("ok"):
        if request.endpoint in ("carregar", "progresso", "tabela", "gerar_pdf"):
            return _texto("Sessão expirada: entre de novo.", 401)
        return redirect(url_for("login", proximo=request.full_path))


@app.route("/login", methods=["GET", "POST"])
def login():
    proximo = request.values.get("proximo") or "/"
    if not proximo.startswith("/") or proximo.startswith("//") or "\\" in proximo:
        proximo = "/"
    if request.method == "POST":
        if not APP_PASSWORD:
            return render_template("login.html", erro="APP_PASSWORD não configurada no servidor.", proximo=proximo), 500
        if hmac.compare_digest(request.form.get("senha", "").encode(), APP_PASSWORD.encode()):
            session["ok"] = True
            session.permanent = True
            return redirect(proximo)
        time.sleep(1)  # freia tentativas
        return render_template("login.html", erro="Senha incorreta.", proximo=proximo), 401
    return render_template("login.html", erro=None, proximo=proximo)


@app.get("/sair")
def sair():
    session.clear()
    return redirect(url_for("login"))


@app.get("/")
def inicio():
    # sem período na URL não consulta sozinho: cada consulta é longa e bloqueia as outras (TRAVA)
    ini = date(2026, 1, 1)
    fim = min(date.fromisoformat(os.getenv("VENC_FIM") or date.today().isoformat()), limite_6_meses(ini))
    return render_template("index.html", ini=request.args.get("venc_ini", ini.isoformat()),
                           fim=request.args.get("venc_fim", fim.isoformat()),
                           consultar="venc_ini" in request.args, baixar=request.args.get("baixar") == "1")


@app.post("/carregar")
def carregar():
    try:
        d_ini = date.fromisoformat(request.form.get("venc_ini", ""))
        d_fim = date.fromisoformat(request.form.get("venc_fim", ""))
    except ValueError:
        return _texto("Datas inválidas.", 400)
    if d_fim < d_ini:
        return _texto("A data final é antes da inicial.", 400)
    if d_fim > limite_6_meses(d_ini):
        return _texto(f"O período deve ter no máximo 6 meses (até {limite_6_meses(d_ini):%d/%m/%Y}).", 400)
    for jid, j in list(JOBS.items()):  # recarregou a página no meio da busca: segue a que já está rodando
        if (j["ini"], j["fim"]) == (d_ini, d_fim) and not j["pronto"] and not j["erro"]:
            return jsonify(job=jid)
    if not TRAVA.acquire(blocking=False):
        return _texto("Já tem um relatório sendo gerado. Aguarde terminar e tente de novo.", 409)

    agora = time.time()
    for k in [k for k, j in list(JOBS.items()) if agora - j["t"] > JOB_TTL]:
        del JOBS[k]
    jid = secrets.token_urlsafe(16)
    JOBS[jid] = job = {"t": agora, "ini": d_ini, "fim": d_fim, "pct": 0, "etapa": "Iniciando…",
                       "pronto": False, "erro": None,
                       "pdf": None, "pdf_gerando": False, "pdf_pag": 0, "pdf_erro": None}
    threading.Thread(target=_rodar, args=(job, d_ini, d_fim), daemon=True).start()
    return jsonify(job=jid)


@app.get("/progresso/<jid>")
def progresso(jid: str):
    j = JOBS.get(jid) or abort(404)
    return jsonify(pct=round(j["pct"], 1), etapa=j["etapa"], pronto=j["pronto"], erro=j["erro"],
                   pdf_pag=j["pdf_pag"], pdf_pronto=bool(j["pdf"]), pdf_erro=j["pdf_erro"])


@app.get("/tabela/<jid>")
def tabela(jid: str):
    j = JOBS.get(jid)
    if not j or not j.get("r"):
        abort(404)
    return render_template("_tabela.html", r=j["r"], job=jid, ini=j["ini"], fim=j["fim"], SIT=PDF.SIT_CURTA)


@app.post("/pdf/<jid>/gerar")
def gerar_pdf(jid: str):
    j = JOBS.get(jid)
    if not j or not j.get("r"):
        abort(404)
    if j["pdf"] or j["pdf_gerando"]:  # outra aba já pediu
        return jsonify(ok=True)
    if not TRAVA.acquire(blocking=False):
        return _texto("Já tem um relatório sendo gerado. Aguarde terminar e tente de novo.", 409)
    j.update(pdf_gerando=True, pdf_pag=0, pdf_erro=None)
    threading.Thread(target=_rodar_pdf, args=(j,), daemon=True).start()
    return jsonify(ok=True)


@app.get("/pdf/<jid>")
def baixar(jid: str):
    j = JOBS.get(jid)
    if not j or not j.get("pdf"):
        # consulta expirou ou o servidor reiniciou (o plano free dorme): refaz com a barra e baixa no fim
        if "venc_ini" not in request.args:
            abort(404)
        return redirect(url_for("inicio", venc_ini=request.args["venc_ini"],
                                venc_fim=request.args.get("venc_fim"), baixar=1))
    return send_file(io.BytesIO(j["pdf"]), mimetype="application/pdf", as_attachment=True,
                     download_name=j["nome"])


if __name__ == "__main__":
    app.run(debug=False, port=int(os.getenv("PORT", "5000")))
