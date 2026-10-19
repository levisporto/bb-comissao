#!/usr/bin/env python3
"""
Web app: senha -> PDF "faturas_total_e_baixas" (mesmo PDF do relatorio_faturas_pdf.py)

Local:   APP_PASSWORD=xxx python app.py        -> http://localhost:5000
Render:  gunicorn app:app   (ver render.yaml; precisa ser 1 worker: os jobs ficam na memória)

Variáveis: APP_PASSWORD + as TOTVS_* do .env (no Render, em Environment).

Fluxo: POST / valida e dispara a geração numa thread -> o navegador consulta
/progresso/<job> a cada segundo (barra de %) -> baixa em /pdf/<job>.
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

from flask import Flask, abort, jsonify, render_template_string, request, send_file

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
TRAVA = threading.Lock()  # a geração mexe em variáveis globais dos scripts: uma por vez
JOBS: dict[str, dict] = {}
JOB_TTL = 30 * 60         # PDFs ficam disponíveis por 30 min

app = Flask(__name__)

PAGINA = """<!doctype html>
<html lang="pt-br">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Faturas - Total e Baixas</title>
<style>
  body { font-family: system-ui, sans-serif; background: #f4f4f4; margin: 0;
         display: flex; min-height: 100vh; align-items: center; justify-content: center; }
  form { background: #fff; padding: 28px; border-radius: 10px; width: 100%; max-width: 340px;
         box-shadow: 0 2px 10px rgba(0,0,0,.08); margin: 16px; }
  h1 { font-size: 18px; margin: 0 0 18px; }
  label { display: block; font-size: 13px; margin: 12px 0 4px; color: #444; }
  input { width: 100%; box-sizing: border-box; padding: 9px; font-size: 15px;
          border: 1px solid #ccc; border-radius: 6px; }
  .datas { display: flex; gap: 8px; } .datas div { flex: 1; }
  button { margin-top: 20px; width: 100%; padding: 11px; font-size: 15px; border: 0;
           border-radius: 6px; background: #1f5c3a; color: #fff; cursor: pointer; }
  button:disabled { opacity: .6; cursor: default; }
  .erro { background: #fde2e2; color: #8a1f1f; padding: 9px; border-radius: 6px; font-size: 13px;
          margin-bottom: 8px; white-space: pre-wrap; }
  .dica { font-size: 12px; color: #777; margin-top: 10px; }
  #de_novo { background: #fff; color: #1f5c3a; border: 1px solid #1f5c3a; margin-top: 10px; }
  #prog { margin-top: 16px; }
  .linha { display: flex; align-items: center; gap: 8px; font-size: 13px; color: #444; }
  .linha span:last-child { margin-left: auto; font-variant-numeric: tabular-nums; font-weight: 600; }
  .spin { width: 16px; height: 16px; border: 2px solid #c9dccf; border-top-color: #1f5c3a;
          border-radius: 50%; animation: gira .8s linear infinite; flex: none; }
  @keyframes gira { to { transform: rotate(360deg); } }
  .barra { height: 8px; background: #e3ede6; border-radius: 4px; overflow: hidden; margin-top: 8px; }
  .barra div { height: 100%; width: 0; background: #1f5c3a; transition: width .4s; }
  [hidden] { display: none !important; }
</style>
</head>
<body>
<form id="f">
  <h1>🐵 Bebetenkite - Faturas do atacado<br><small>total vendido e baixas</small></h1>
  <div id="erro" class="erro" hidden></div>
  <label for="senha">Senha</label>
  <input id="senha" name="senha" type="password" required autofocus>
  <div class="datas">
    <div><label for="ini">Vencimento de</label><input id="ini" name="venc_ini" type="date" value="{{ ini }}" required></div>
    <div><label for="fim">até</label><input id="fim" name="venc_fim" type="date" value="{{ fim }}" required></div>
  </div>
  <button id="b" type="submit">Gerar PDF</button>
  <div id="prog" hidden>
    <div class="linha"><div class="spin" id="spin"></div><span id="etapa">Iniciando…</span><span id="pct">0%</span></div>
    <div class="barra"><div id="barra"></div></div>
  </div>
  <button id="de_novo" type="button" hidden>Baixar novamente</button>
  <div class="dica" id="dica">Período de no máximo 6 meses (limite do TOTVS). O download começa sozinho quando o PDF ficar pronto.</div>
</form>
<script>
// TOTVS só aceita períodos de até 6 meses: "até" fica entre "de" e "de" + 6 meses - 1 dia
function limite(s) {
  const [y, m, d] = s.split("-").map(Number);
  const ultimo = new Date(Date.UTC(y, m + 6, 0)).getUTCDate();  // último dia do mês alvo
  return new Date(Date.UTC(y, m - 1 + 6, Math.min(d, ultimo)) - 864e5).toISOString().slice(0, 10);
}
function ajustarFim() {
  if (!ini.value) return;
  fim.min = ini.value; fim.max = limite(ini.value);
  if (fim.value > fim.max) fim.value = fim.max;
  if (fim.value && fim.value < fim.min) fim.value = fim.min;
}
ini.addEventListener("change", ajustarFim);
ajustarFim();
</script>
<script>
// o PDF fica guardado no navegador: "Baixar novamente" não consulta a API de novo
let pdf = null, nome = "faturas.pdf";
const espera = (ms) => new Promise((r) => setTimeout(r, ms));
function baixar() {
  const a = document.createElement("a");
  a.href = pdf; a.download = nome; document.body.appendChild(a); a.click(); a.remove();
}
function mostrar(p, txt) {
  barra.style.width = p + "%"; pct.textContent = Math.floor(p) + "%"; etapa.textContent = txt;
}
de_novo.onclick = baixar;
f.onsubmit = async (ev) => {
  ev.preventDefault();
  erro.hidden = true;
  if (fim.value < ini.value || fim.value > limite(ini.value)) {
    erro.textContent = "O período deve ter no máximo 6 meses (até " +
      limite(ini.value).split("-").reverse().join("/") + ").";
    erro.hidden = false;
    return;
  }
  b.disabled = true;
  prog.hidden = false; spin.hidden = false; mostrar(0, "Iniciando…");
  try {
    let r = await fetch("/", { method: "POST", body: new FormData(f) });
    if (!r.ok) throw new Error(await r.text());
    const { job } = await r.json();
    for (;;) {
      await espera(1000);
      r = await fetch("/progresso/" + job);
      if (!r.ok) throw new Error("Perdi o acompanhamento da geração. Tente de novo.");
      const s = await r.json();
      if (s.erro) throw new Error(s.erro);
      mostrar(s.pct, s.etapa);
      if (s.pronto) break;
    }
    r = await fetch("/pdf/" + job);
    if (!r.ok) throw new Error("Não consegui baixar o PDF. Tente de novo.");
    nome = (/filename="?([^";]+)"?/.exec(r.headers.get("Content-Disposition") || "") || [, nome])[1];
    if (pdf) URL.revokeObjectURL(pdf);
    pdf = URL.createObjectURL(await r.blob());
    baixar();
    mostrar(100, "Pronto");
    de_novo.hidden = false;
    dica.textContent = "Pronto: " + nome;
  } catch (e) {
    prog.hidden = true;
    erro.textContent = e.message || "Falha de conexão.";
    erro.hidden = false;
  } finally {
    spin.hidden = true; b.disabled = false;
  }
};
</script>
</body>
</html>"""


# ----------------------------------------------------------------------------
# Progresso: etapas com faixas de % (títulos 2-45, notas 45-75, clientes 75-92, PDF 92-100)
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
    ATUAL["pct"], ATUAL["etapa"] = 92, "Montando o PDF"
    return resultado


R.buscar_pessoas = _buscar_pessoas
R.PROGRESSO = _progresso


def _rodar(job: dict, d_ini: date, d_fim: date) -> None:
    global ATUAL
    try:
        ATUAL = job
        job["etapa"] = "Conectando ao TOTVS"
        S.VENC_INI, S.VENC_FIM = d_ini.isoformat(), d_fim.isoformat()
        agora = datetime.now()
        buf = io.BytesIO()
        PDF.gerar(buf, agora)
        job["pdf"], job["nome"] = buf.getvalue(), PDF.nome_arquivo(agora)
        job["pct"], job["etapa"], job["pronto"] = 100, "Pronto", True
    except SystemExit as e:  # os scripts usam sys.exit("mensagem") para erros esperados
        job["erro"] = str(e)
    except Exception as e:  # noqa: BLE001
        traceback.print_exc()
        job["erro"] = f"Erro ao gerar o PDF: {e}"
    finally:
        ATUAL = {}
        TRAVA.release()


limite_6_meses = PDF.RF.limite_6_meses


def _texto(msg: str, status: int):
    return msg, status, {"Content-Type": "text/plain; charset=utf-8"}


# ----------------------------------------------------------------------------
# Rotas
# ----------------------------------------------------------------------------

@app.get("/")
def inicio():
    ini = date(2026, 1, 1)
    fim = min(date.fromisoformat(os.getenv("VENC_FIM") or date.today().isoformat()), limite_6_meses(ini))
    return render_template_string(PAGINA, ini=ini.isoformat(), fim=fim.isoformat())


@app.post("/")
def iniciar():
    senha = request.form.get("senha", "")
    if not APP_PASSWORD:
        return _texto("APP_PASSWORD não configurada no servidor.", 500)
    if not hmac.compare_digest(senha.encode(), APP_PASSWORD.encode()):
        time.sleep(1)  # freia tentativas
        return _texto("Senha incorreta.", 401)
    try:
        d_ini = date.fromisoformat(request.form.get("venc_ini", ""))
        d_fim = date.fromisoformat(request.form.get("venc_fim", ""))
    except ValueError:
        return _texto("Datas inválidas.", 400)
    if d_fim < d_ini:
        return _texto("A data final é antes da inicial.", 400)
    if d_fim > limite_6_meses(d_ini):
        return _texto(f"O período deve ter no máximo 6 meses (até {limite_6_meses(d_ini):%d/%m/%Y}).", 400)
    if not TRAVA.acquire(blocking=False):
        return _texto("Já tem um PDF sendo gerado. Aguarde terminar e tente de novo.", 409)

    agora = time.time()
    for k in [k for k, j in JOBS.items() if agora - j["t"] > JOB_TTL]:
        del JOBS[k]
    jid = secrets.token_urlsafe(16)  # o id é a "chave" do download
    JOBS[jid] = job = {"t": agora, "pct": 0, "etapa": "Iniciando…", "pronto": False, "erro": None}
    threading.Thread(target=_rodar, args=(job, d_ini, d_fim), daemon=True).start()
    return jsonify(job=jid)


@app.get("/progresso/<jid>")
def progresso(jid: str):
    j = JOBS.get(jid) or abort(404)
    return jsonify(pct=round(j["pct"], 1), etapa=j["etapa"], pronto=j["pronto"], erro=j["erro"])


@app.get("/pdf/<jid>")
def baixar(jid: str):
    j = JOBS.get(jid)
    if not j or not j.get("pdf"):
        abort(404)
    return send_file(io.BytesIO(j["pdf"]), mimetype="application/pdf", as_attachment=True,
                     download_name=j["nome"])


if __name__ == "__main__":
    app.run(debug=False, port=int(os.getenv("PORT", "5000")))
