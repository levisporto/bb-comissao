#!/usr/bin/env python3
"""
Sonda da API TOTVS Moda - Controle de Vendas Atacado / Comissao de Corretores
=============================================================================

Objetivo: descobrir QUAIS informacoes a API devolve, para saber se ela substitui
o CSV do relatorio 0433 (Controle Atacado) e se resolve:
  - numero da Transacao (que vem zerado em baixas parciais)
  - representante/corretor vinculado ao cliente
  - data e tipo de baixa (pagamento)
  - valor total do fechamento / valor da venda (romaneio)

O que o script faz:
  1. Autentica (token OAuth password grant).
  2. (Opcional) Le os swagger.json informados em SWAGGER_URLS e lista os
     endpoints e campos que interessam (transacao, representante, baixa...).
  3. Chama cada endpoint configurado em ENDPOINTS (abaixo), com 1-2 paginas.
     - Se der erro 400 com "expand", tenta de novo sem expand.
     - Salva a resposta bruta em saida_api/<endpoint>.json
  4. Faz o inventario de TODOS os campos devolvidos (caminho, tipo, % preenchido,
     exemplos) -> saida_api/inventario_campos.csv
  5. Compara com as colunas do CSV atual (Corretor, Cliente, Nr. Fat,
     Transacao, Valor, Tp. Baixa...) -> saida_api/relatorio.md
  6. Conta quantas transacoes vem com valor 0 (o problema da planilha).

Uso:
  pip install requests
  copie .env.example para .env e preencha
  python sonda_api_totvs.py                 # roda tudo
  python sonda_api_totvs.py --so-swagger    # so le os swaggers
  python sonda_api_totvs.py --endpoint ar_documentos_vencimento   # so um

IMPORTANTE: os caminhos e nomes de filtros marcados "CONFIRMAR NO SWAGGER"
seguem o padrao da API v2 do TOTVS Moda, mas precisam ser conferidos.
O proprio erro 400 da API costuma dizer qual campo esta errado - o
relatorio guarda essa mensagem.
"""

from __future__ import annotations

import argparse
import csv
import json
import os
import re
import sys
import time
from collections import defaultdict
from datetime import datetime
from pathlib import Path

try:
    import requests
except ImportError:
    sys.exit("Instale o requests:  pip install requests")


# ----------------------------------------------------------------------------
# Configuracao (.env)
# ----------------------------------------------------------------------------

def carregar_env(caminho: Path) -> None:
    """Le um .env simples (CHAVE=valor) sem depender de python-dotenv."""
    if not caminho.exists():
        return
    for linha in caminho.read_text(encoding="utf-8").splitlines():
        linha = linha.strip()
        if not linha or linha.startswith("#") or "=" not in linha:
            continue
        chave, valor = linha.split("=", 1)
        os.environ.setdefault(chave.strip(), valor.strip().strip('"').strip("'"))


PASTA = Path(__file__).resolve().parent
carregar_env(PASTA / ".env")

BASE_URL = os.getenv("TOTVS_BASE_URL", "https://apitotvsmoda.bhan.com.br").rstrip("/")
TOKEN_PATH = os.getenv("TOTVS_TOKEN_PATH", "/api/totvsmoda/authorization/v2/token")
CLIENT_ID = os.getenv("TOTVS_CLIENT_ID", "")
CLIENT_SECRET = os.getenv("TOTVS_CLIENT_SECRET", "")
USERNAME = os.getenv("TOTVS_USERNAME", "")
PASSWORD = os.getenv("TOTVS_PASSWORD", "")

# "Empresa" no CSV = 2
BRANCHES = [int(x) for x in os.getenv("TOTVS_BRANCH_CODES", "2").split(",") if x.strip()]

# Periodo de teste: o fechamento de 05/10/2026 (emissoes de ~31/08 a 30/09)
VENC_INI = os.getenv("VENC_INI", "2026-10-05")
VENC_FIM = os.getenv("VENC_FIM", "2026-10-05")
EMISSAO_INI = os.getenv("EMISSAO_INI", "2026-08-25")
EMISSAO_FIM = os.getenv("EMISSAO_FIM", "2026-09-30")

# Clientes conhecidos da planilha, para comparar linha a linha
#   104691 = H T DE ARAUJO (Nr. Fat 36628 tem uma linha com transacao 0)
#   112507 = ANA PAULA F OLIVEIRA (cliente da ANTONIA ALINE)
CLIENTES_TESTE = [int(x) for x in os.getenv("CLIENTES_TESTE", "104691,112507").split(",") if x.strip()]

PAGE_SIZE = int(os.getenv("PAGE_SIZE", "100"))
MAX_PAGINAS = int(os.getenv("MAX_PAGINAS", "2"))  # sonda, nao carga completa
TIMEOUT = int(os.getenv("TIMEOUT", "60"))

# URLs de swagger.json (separadas por virgula) - opcional
SWAGGER_URLS = [u.strip() for u in os.getenv("SWAGGER_URLS", "").split(",") if u.strip()]

SAIDA = PASTA / "saida_api"

# Transacoes da planilha que estavam zeradas (recuperadas pela NF)
TRANSACOES_TESTE = [int(x) for x in os.getenv("TRANSACOES_TESTE", "905453,904885,905181").split(",") if x.strip()]


def data_iso(d: str, fim: bool = False) -> str:
    """'2026-10-05' -> '2026-10-05T00:00:00' (ou 23:59:59 se fim)."""
    return f"{d}T23:59:59" if fim else f"{d}T00:00:00"


# ----------------------------------------------------------------------------
# Endpoints a sondar  (CONFIRMAR NO SWAGGER: path, nomes de filtro, expand)
# ----------------------------------------------------------------------------
# Padrao das buscas v2 do TOTVS Moda:
#   POST .../search   body = {"filter": {...}, "expand": "a,b", "page": 1, "pageSize": 100}
#   resposta = {"count", "totalPages", "hasNext", "totalItems", "items": [...]}

ENDPOINTS: dict[str, dict] = {
    # 1) Contas a receber pelo VENCIMENTO -> equivale ao CSV 0433
    "ar_documentos_vencimento": {
        "descricao": "Contas a receber com vencimento no fechamento (substituto do CSV 0433)",
        "metodo": "POST",
        "path": "/api/totvsmoda/accounts-receivable/v2/documents/search",
        "body": {
            "filter": {
                "branchCodeList": BRANCHES,
                "startExpiredDate": data_iso(VENC_INI),
                "endExpiredDate": data_iso(VENC_FIM, fim=True),
            },
            # expand tenta trazer fatura, valores calculados e comissoes
            "expand": "invoice,calculateValue,commissions",
        },
    },
    # 2) Contas a receber dos clientes de teste (comparar com a planilha)
    "ar_documentos_clientes_teste": {
        "descricao": "Contas a receber de clientes conhecidos da planilha (comparacao linha a linha)",
        "metodo": "POST",
        "path": "/api/totvsmoda/accounts-receivable/v2/documents/search",
        "body": {
            "filter": {
                "branchCodeList": BRANCHES,
                "customerCodeList": CLIENTES_TESTE,
                "startExpiredDate": data_iso(VENC_INI),
                "endExpiredDate": data_iso(VENC_FIM, fim=True),
            },
            "expand": "invoice,calculateValue,commissions",
        },
    },
    # 3) Notas fiscais de venda do periodo -> possivel origem do "romaneio"/valor da venda
    "fiscal_notas_periodo": {
        "descricao": "Notas fiscais emitidas no periodo (valor da venda / romaneio?)",
        "metodo": "POST",
        "path": "/api/totvsmoda/fiscal/v2/invoices/search",
        "body": {
            "filter": {
                "branchCodeList": BRANCHES,
                "personCodeList": CLIENTES_TESTE,
                "startIssueDate": data_iso(EMISSAO_INI),
                "endIssueDate": data_iso(EMISSAO_FIM, fim=True),
            },
            "expand": "items,payments,representatives",
        },
    },
    # 4) Pedidos de venda (alternativa para o valor da venda e representante)
    "pedidos_venda_periodo": {
        "descricao": "Pedidos de venda do periodo (representante e valor do pedido)",
        "metodo": "POST",
        "path": "/api/totvsmoda/sales-order/v2/orders/search",
        "body": {
            "filter": {
                "branchCodeList": BRANCHES,
                "customerCodeList": CLIENTES_TESTE,
                "startOrderDate": data_iso(EMISSAO_INI),
                "endOrderDate": data_iso(EMISSAO_FIM, fim=True),
            },
            "expand": "representatives,items",
        },
    },
    # 5) Cadastro do cliente PJ -> vinculo cliente x representante
    "pessoa_juridica_clientes": {
        "descricao": "Cadastro PJ dos clientes de teste (vinculo com representante)",
        "metodo": "POST",
        "path": "/api/totvsmoda/person/v2/legal-entities/search",
        "body": {
            "filter": {"personCodeList": CLIENTES_TESTE},
            "expand": "representatives,customers",
        },
    },
    # ---- Novos (lista do TDN) - caminhos e filtros sao PALPITE; o erro 400 diz o que falta ----
    # 7) Fechamento de comissao (Accounts Payable) - o TOTVS ja guarda comissao paga?
    "comissoes_pagas": {
        "descricao": "Fechamento de comissão (accounts-payable comissions-paid/search)",
        "metodo": "POST",
        "path": "/api/totvsmoda/accounts-payable/v2/comissions-paid/search",
        # swagger: filter.closingCompanyCode* ; empresa 2 voltou vazia -> testar a empresa 1
        "bodies": [
            {"filter": {"closingCompanyCode": int(os.getenv("EMPRESA_FECHAMENTO_COMISSAO", "1"))},
             "expand": "accountMovements"},
        ],
    },
    # 7b) Comissao paga via contas a pagar: duplicatas com documentType 3 = Comissao
    "ap_duplicatas_comissao": {
        "descricao": "Contas a pagar - duplicatas do tipo 3 (Comissão) pagas a corretores em 2026",
        "metodo": "POST",
        "path": "/api/totvsmoda/accounts-payable/v2/duplicates/search",
        "body": {"filter": {"branchCodeList": [int(x) for x in os.getenv("EMPRESAS_PAGAR", "1,2").split(",")],
                            "documentTypeList": [3],
                            "startIssueDate": data_iso("2026-01-01"),
                            "endIssueDate": data_iso(VENC_FIM, fim=True)},
                 "expand": "classifications"},
    },
    # 8) Segunda busca de contas a receber - pode trazer campos diferentes da documents/search
    "ar_receivable_search": {
        "descricao": "Contas a receber - receivable/search (comparar campos com documents/search)",
        "metodo": "POST",
        "path": "/api/totvsmoda/accounts-receivable/v2/receivable/search",
        "body": {"filter": {"branchCodeList": BRANCHES, "customerCodeList": CLIENTES_TESTE,
                            "startExpiredDate": data_iso(VENC_INI),
                            "endExpiredDate": data_iso(VENC_FIM, fim=True)}},
    },
    # 9) Transacoes (General) - a "Transacao" do CSV; pode ter representante/vendedor
    "transacoes": {
        "descricao": "Consulta dados de transação (general transactions) - transações conhecidas da planilha",
        "metodo": "POST",
        "path": "/api/totvsmoda/general/v2/transactions/search",
        # erro anterior: "Empty field BranchCodeList" com o filtro dentro de "filter" -> tentar solto
        "bodies": [
            {"branchCodeList": BRANCHES, "transactionCodeList": TRANSACOES_TESTE,
             "expand": "representatives,sellers,items"},
            {"branchCodeList": BRANCHES, "transactionCodeList": TRANSACOES_TESTE},
            {"branchCodeList": BRANCHES, "transactionCode": TRANSACOES_TESTE[0]},
            {"branchCodeList": BRANCHES, "startTransactionDate": data_iso("2026-09-02"),
             "endTransactionDate": data_iso("2026-09-02", fim=True)},
        ],
    },
    # 10) Representantes da movimentacao fiscal (Analytics) - vendas por representante prontas?
    "analytics_representante": {
        "descricao": "Analytics - representantes da movimentação fiscal",
        "metodo": "POST",
        "path": "/api/totvsmoda/analytics/v2/representative-fiscal-movement/search",
        "body": {"filter": {"branchCodeList": BRANCHES,
                            "startMovementDate": data_iso(EMISSAO_INI),
                            "endMovementDate": data_iso(EMISSAO_FIM, fim=True)}},
    },
    # 11) Vendas e devolucoes por empresa (Analytics)
    "analytics_vendas_devolucoes": {
        "descricao": "Analytics - vendas e devoluções de vendas por empresa (branch-sale)",
        "metodo": "POST",
        "path": "/api/totvsmoda/analytics/v2/branch-sale/search",
        "body": {"filter": {"branchCodeList": BRANCHES,
                            "startMovementDate": data_iso(EMISSAO_INI),
                            "endMovementDate": data_iso(EMISSAO_FIM, fim=True)}},
    },
    # 12) Devolucoes (General)
    "devolucoes": {
        "descricao": "Consulta dados de devolução (general devolutions/search)",
        "metodo": "POST",
        "path": "/api/totvsmoda/general/v2/devolutions/search",
        "body": {"filter": {"branchCodeList": BRANCHES, "customerCodeList": CLIENTES_TESTE,
                            "startDate": data_iso(EMISSAO_INI), "endDate": data_iso(VENC_FIM, fim=True)}},
    },
    # 13) Cadastro de representantes (Person)
    "representantes": {
        "descricao": "Consultar dados de representante (person representatives)",
        "metodo": "POST",
        "path": "/api/totvsmoda/person/v2/representatives/search",
        "body": {"filter": {}},
    },
    # 6) Cadastro do cliente PF (a maioria dos clientes do atacado tem CPF)
    "pessoa_fisica_clientes": {
        "descricao": "Cadastro PF dos clientes de teste (vinculo com representante)",
        "metodo": "POST",
        "path": "/api/totvsmoda/person/v2/individuals/search",
        "body": {
            "filter": {"personCodeList": CLIENTES_TESTE},
            "expand": "representatives,customers",
        },
    },
}


# ----------------------------------------------------------------------------
# Colunas do CSV atual x palavras-chave para achar o campo equivalente na API
# ----------------------------------------------------------------------------

COLUNAS_CSV = {
    "Corretor / Representante": ["represent", "commission", "agent", "seller", "salesperson"],
    "CD cliente": ["customercode", "personcode", "clientcode"],
    "Cliente": ["customername", "personname", "name"],
    "Cpfcnpj": ["cpf", "cnpj", "federal", "taxid"],
    "Vencimento": ["expired", "duedate", "expiration", "maturity"],
    "Emissao": ["issue"],
    "Prz (prazo)": ["term", "deadline", "days"],
    "Nr. Fat": ["receivablecode", "invoicecode", "documentnumber", "duplicate", "installment"],
    "Dt. Baixa": ["settlement", "payment", "paid", "liquidat", "discharge", "write"],
    "Transacao": ["transaction"],
    "Valor": ["value", "amount", "total"],
    "Tp. Baixa": ["settlementtype", "dischargetype", "paymenttype", "status", "type"],
    "Romaneio / venda (novo)": ["romaneio", "packing", "order", "invoice", "fiscal", "sale"],
}


# ----------------------------------------------------------------------------
# HTTP
# ----------------------------------------------------------------------------

class Api:
    def __init__(self) -> None:
        self.s = requests.Session()
        self.token: str | None = None

    def autenticar(self) -> None:
        faltando = [n for n, v in {
            "TOTVS_CLIENT_ID": CLIENT_ID, "TOTVS_CLIENT_SECRET": CLIENT_SECRET,
            "TOTVS_USERNAME": USERNAME, "TOTVS_PASSWORD": PASSWORD}.items() if not v]
        if faltando:
            sys.exit(f"Preencha no .env: {', '.join(faltando)}")
        r = self.s.post(
            BASE_URL + TOKEN_PATH,
            data={"grant_type": "password", "client_id": CLIENT_ID,
                  "client_secret": CLIENT_SECRET, "username": USERNAME, "password": PASSWORD},
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            timeout=TIMEOUT,
        )
        if r.status_code != 200:
            sys.exit(f"Falha na autenticacao ({r.status_code}): {r.text[:500]}")
        self.token = r.json().get("access_token")
        if not self.token:
            sys.exit(f"Resposta de token sem access_token: {list(r.json().keys())}")
        self.s.headers.update({"Authorization": f"Bearer {self.token}"})
        print("✔ Autenticado")

    def chamar(self, metodo: str, path: str, body: dict | None) -> requests.Response:
        url = path if path.startswith("http") else BASE_URL + path
        return self.s.request(metodo, url, json=body, timeout=TIMEOUT)


# ----------------------------------------------------------------------------
# Inventario de campos
# ----------------------------------------------------------------------------

class Inventario:
    """Percorre o JSON e registra cada caminho de campo, tipo e exemplos."""

    def __init__(self) -> None:
        self.campos: dict[str, dict] = defaultdict(
            lambda: {"presente": 0, "nao_nulo": 0, "tipos": set(), "exemplos": [], "zeros": 0})

    def registrar(self, obj, caminho: str = "") -> None:
        if isinstance(obj, dict):
            for k, v in obj.items():
                self.registrar(v, f"{caminho}.{k}" if caminho else k)
        elif isinstance(obj, list):
            if not obj:
                c = self.campos[caminho + "[]"]
                c["presente"] += 1
                c["tipos"].add("lista_vazia")
            for item in obj:
                self.registrar(item, caminho + "[]")
        else:
            c = self.campos[caminho]
            c["presente"] += 1
            c["tipos"].add(type(obj).__name__)
            if obj not in (None, ""):
                c["nao_nulo"] += 1
                if obj == 0 or obj == "0":
                    c["zeros"] += 1
                if len(c["exemplos"]) < 3 and obj not in c["exemplos"]:
                    c["exemplos"].append(obj)


def extrair_itens(dados) -> list:
    if isinstance(dados, list):
        return dados
    if isinstance(dados, dict):
        for chave in ("items", "data", "results", "content"):
            if isinstance(dados.get(chave), list):
                return dados[chave]
        return [dados]
    return []


# ----------------------------------------------------------------------------
# Sondagem de endpoints
# ----------------------------------------------------------------------------

def sondar_endpoint(api: Api, nome: str, cfg: dict) -> dict:
    print(f"\n▶ {nome}: {cfg['descricao']}")
    resultado = {"nome": nome, "descricao": cfg["descricao"], "path": cfg["path"],
                 "status": None, "erro": None, "expand_usado": None,
                 "total_itens_api": None, "itens_lidos": 0, "inventario": None,
                 "tentativas": []}

    if cfg.get("bodies"):  # varios formatos de body: tenta ate um funcionar
        variantes = [json.loads(json.dumps(b)) for b in cfg["bodies"]]
    else:
        body_base = json.loads(json.dumps(cfg.get("body") or {}))  # copia
        variantes = [body_base]
        if body_base.get("expand"):
            variantes.append({k: v for k, v in body_base.items() if k != "expand"})

    for body in variantes:
        itens, paginas_brutas = [], []
        ok = True
        for pagina in range(1, MAX_PAGINAS + 1):
            if cfg["metodo"] == "POST":
                body_pag = {**body, "page": pagina, "pageSize": PAGE_SIZE}
            else:
                body_pag = None
            t0 = time.time()
            try:
                r = api.chamar(cfg["metodo"], cfg["path"], body_pag)
            except requests.RequestException as e:
                resultado["erro"] = f"Falha de conexao: {e}"
                ok = False
                break
            dt = time.time() - t0
            resultado["tentativas"].append({
                "expand": body.get("expand"), "pagina": pagina,
                "status": r.status_code, "segundos": round(dt, 2),
                "resposta_erro": None if r.ok else r.text[:800]})
            print(f"   pagina {pagina} expand={body.get('expand')!r} -> HTTP {r.status_code} ({dt:.1f}s)")
            if not r.ok:
                resultado["status"] = r.status_code
                resultado["erro"] = r.text[:800]
                ok = False
                break
            try:
                dados = r.json()
            except ValueError:
                resultado["status"] = r.status_code
                resultado["erro"] = "Resposta nao e JSON: " + r.text[:300]
                ok = False
                break
            paginas_brutas.append(dados)
            itens.extend(extrair_itens(dados))
            if isinstance(dados, dict):
                # "count" e o tamanho da pagina; o total real e totalItems
                resultado["total_itens_api"] = dados.get("totalItems") or (
                    f"{dados.get('totalPages')} paginas" if dados.get("totalPages") else None)
                if not dados.get("hasNext"):
                    break
            else:
                break

        if ok:
            resultado["status"] = 200
            resultado["erro"] = None
            resultado["expand_usado"] = body.get("expand")
            resultado["itens_lidos"] = len(itens)
            inv = Inventario()
            for item in itens:
                inv.registrar(item)
            resultado["inventario"] = inv
            resultado["amostra"] = itens[:200]
            resultado["body_que_funcionou"] = body
            SAIDA.mkdir(exist_ok=True)
            (SAIDA / f"{nome}.json").write_text(
                json.dumps({"request_body": body, "paginas": paginas_brutas},
                           ensure_ascii=False, indent=2, default=str),
                encoding="utf-8")
            print(f"   ✔ {len(itens)} itens lidos (total na API: {resultado['total_itens_api']}), "
                  f"{len(inv.campos)} campos distintos")
            return resultado
        # tenta a proxima variante de body (sem expand / outro formato)
        if body is not variantes[-1] and resultado["status"] in (400, 422, 500):
            print("   ↻ erro; tentando o proximo formato de body")
            continue
        break

    print(f"   ✘ erro: {str(resultado['erro'])[:200]}")
    return resultado


# ----------------------------------------------------------------------------
# Swagger
# ----------------------------------------------------------------------------

PALAVRAS_INTERESSE = ["transaction", "represent", "commission", "settlement", "receivable",
                      "invoice", "payment", "expired", "discharge", "romaneio", "order"]


def ler_swaggers(api: Api) -> list[dict]:
    achados = []
    for url in SWAGGER_URLS:
        print(f"\n▶ swagger: {url}")
        try:
            r = api.s.get(url, timeout=TIMEOUT)
            spec = r.json()
        except Exception as e:  # noqa: BLE001
            print(f"   ✘ nao consegui ler: {e}")
            achados.append({"url": url, "erro": str(e)})
            continue
        paths = spec.get("paths", {})
        schemas = spec.get("components", {}).get("schemas", {}) or spec.get("definitions", {})
        lista_paths = [(p, m.upper()) for p, ops in paths.items() for m in ops
                       if m.lower() in ("get", "post", "put", "patch", "delete")]
        campos_interesse = []
        alvos = PALAVRAS_INTERESSE + ["status", "discharge", "documenttype", "operation"]
        for nome_schema, sch in schemas.items():
            for prop, pdef in (sch.get("properties") or {}).items():
                if any(k in prop.lower() for k in alvos):
                    pdef = pdef or {}
                    if "$ref" in pdef:  # enum definido em outro schema
                        ref = schemas.get(pdef["$ref"].split("/")[-1], {})
                        pdef = {**ref, **pdef}
                    extra = []
                    if pdef.get("enum"):
                        extra.append(f"enum={pdef['enum']}")
                    for k in ("x-enumNames", "x-enum-varnames", "x-enum-descriptions"):
                        if pdef.get(k):
                            extra.append(f"{k}={pdef[k]}")
                    if pdef.get("description"):
                        dominio = prop.lower() in ("status", "statuslist", "dischargetype", "dischargetypelist",
                                                   "documenttype", "documenttypelist")
                        d = pdef["description"].replace("\r\n", " ").replace("<br>", " | ")
                        extra.append(f"desc={d if dominio else d[:200]!r}")
                    campos_interesse.append(f"{nome_schema}.{prop}" + (f" → {'; '.join(extra)}" if extra else ""))
        print(f"   ✔ {len(lista_paths)} operacoes, {len(campos_interesse)} campos de interesse")
        achados.append({"url": url, "paths": lista_paths, "campos_interesse": campos_interesse})
        SAIDA.mkdir(exist_ok=True)
        nome_arq = re.sub(r"[^a-zA-Z0-9]+", "_", url)[-80:] + ".json"
        (SAIDA / f"swagger_{nome_arq}").write_text(json.dumps(spec, ensure_ascii=False, indent=2),
                                                    encoding="utf-8")
    return achados


# ----------------------------------------------------------------------------
# Relatorios
# ----------------------------------------------------------------------------

def salvar_inventario_csv(resultados: list[dict]) -> Path:
    SAIDA.mkdir(exist_ok=True)
    arq = SAIDA / "inventario_campos.csv"
    with arq.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["endpoint", "campo", "tipos", "pct_preenchido", "qtd_zeros", "exemplos"])
        for res in resultados:
            inv = res.get("inventario")
            if not inv:
                continue
            n = max(res["itens_lidos"], 1)
            for campo, c in sorted(inv.campos.items()):
                pct = round(100 * c["nao_nulo"] / n, 1) if "[]" not in campo else ""
                w.writerow([res["nome"], campo, ",".join(sorted(c["tipos"])), pct, c["zeros"],
                            " | ".join(str(x)[:60] for x in c["exemplos"])])
    return arq


EXCLUIR = {  # evita falsos positivos (ex.: "Cliente" casando com representativeName)
    "Cliente": ["represent", "branch", "company", "user"],
    "Nr. Fat": ["value", "amount", "date"],
    "Valor": ["code", "type"],
}


def candidatos(inv: Inventario, palavras: list[str], coluna: str = "") -> list[str]:
    """Compara so o nome final do campo (folha), ignorando o caminho."""
    out = []
    excl = EXCLUIR.get(coluna, [])
    for campo in inv.campos:
        folha = campo.replace("[]", "").split(".")[-1].lower().replace("_", "")
        if any(p in folha for p in palavras) and not any(e in folha for e in excl):
            out.append(campo)
    return sorted(out)


def salvar_relatorio_md(resultados: list[dict], swaggers: list[dict]) -> Path:
    L = []
    L.append("# Sonda da API TOTVS Moda - Controle Vendas Atacado")
    L.append(f"\nGerado em {datetime.now():%d/%m/%Y %H:%M} | base: `{BASE_URL}` | "
             f"empresas: {BRANCHES} | vencimento {VENC_INI} a {VENC_FIM} | "
             f"clientes teste: {CLIENTES_TESTE}\n")

    L.append("## 1. Endpoints\n")
    L.append("| Endpoint | Status | Itens lidos | Total na API | Expand aceito | Observacao |")
    L.append("|---|---|---|---|---|---|")
    for r in resultados:
        obs = (r["erro"] or "").replace("\n", " ").replace("|", "/")[:160]
        L.append(f"| `{r['nome']}` | {r['status']} | {r['itens_lidos']} | {r['total_itens_api']} | "
                 f"{r['expand_usado'] or '-'} | {obs} |")

    L.append("\n## 2. Cobertura das colunas do CSV atual\n")
    L.append("Para cada coluna do relatorio 0433, os campos da API que parecem equivalentes "
             "(busca por palavra-chave - conferir manualmente).\n")
    for r in resultados:
        inv = r.get("inventario")
        if not inv:
            continue
        L.append(f"\n### `{r['nome']}`\n")
        L.append("| Coluna CSV | Campos candidatos na API |")
        L.append("|---|---|")
        for col, palavras in COLUNAS_CSV.items():
            cands = candidatos(inv, palavras, col)
            txt = ", ".join(f"`{c}`" for c in cands[:8]) + (" ..." if len(cands) > 8 else "")
            L.append(f"| {col} | {txt or '❌ nenhum'} |")

    L.append("\n## 3. Transacao zerada\n")
    achou = False
    for r in resultados:
        inv = r.get("inventario")
        if not inv:
            continue
        for campo, c in inv.campos.items():
            if "transaction" in campo.lower():
                achou = True
                L.append(f"- `{r['nome']}` → `{campo}`: {c['nao_nulo']} preenchidos, "
                         f"**{c['zeros']} com valor 0**, exemplos {c['exemplos']}")
    if not achou:
        L.append("- Nenhum campo com 'transaction' no nome foi devolvido. "
                 "Verificar no swagger se existe expand ou outro endpoint com a transacao.")

    L.append("\n## 4. Representante / comissao\n")
    achou = False
    for r in resultados:
        inv = r.get("inventario")
        if not inv:
            continue
        caminhos = [c for c in inv.campos if "represent" in c.lower() or "commission" in c.lower()]
        for campo in sorted(caminhos):
            achou = True
            c = inv.campos[campo]
            L.append(f"- `{r['nome']}` → `{campo}`: exemplos {c['exemplos']}")
    if not achou:
        L.append("- Nenhum campo de representante/comissao devolvido.")

    for r in resultados:
        if r.get("body_que_funcionou") is not None and r["nome"] in ("comissoes_pagas", "transacoes"):
            L.append(f"\n### Body que funcionou em `{r['nome']}`\n\n```json\n"
                     f"{json.dumps(r['body_que_funcionou'], ensure_ascii=False, indent=2)}\n```")
        if r["nome"] == "representantes" and r.get("amostra"):
            L.append("\n### Representantes cadastrados oficialmente no TOTVS\n")
            L.append("| code | nome | supervisor | ativo? |\n|---|---|---|---|")
            for it in r["amostra"]:
                L.append(f"| {it.get('code') or it.get('representativeCode')} | {it.get('name')} | "
                         f"{it.get('supervisorName') or ''} | {it.get('isInactive', it.get('status', ''))} |")

    if swaggers:
        L.append("\n## 5. Swagger\n")
        for s in swaggers:
            L.append(f"\n### {s['url']}\n")
            if s.get("erro"):
                L.append(f"Erro: {s['erro']}")
                continue
            L.append("Operacoes de interesse:\n")
            for p, m in s["paths"]:
                if any(k in p.lower() for k in PALAVRAS_INTERESSE):
                    L.append(f"- `{m} {p}`")
            L.append("\nCampos de interesse nos schemas:\n")
            for c in s["campos_interesse"][:200]:
                L.append(f"- `{c}`")

    L.append("\n## 6. Detalhe das tentativas (para depurar erros)\n")
    for r in resultados:
        for t in r["tentativas"]:
            if t["resposta_erro"]:
                L.append(f"- `{r['nome']}` expand={t['expand']!r} pag {t['pagina']} "
                         f"HTTP {t['status']}: `{t['resposta_erro'][:400]}`")

    arq = SAIDA / "relatorio.md"
    SAIDA.mkdir(exist_ok=True)
    arq.write_text("\n".join(L), encoding="utf-8")
    return arq


# ----------------------------------------------------------------------------

def main() -> None:
    ap = argparse.ArgumentParser(description="Sonda da API TOTVS Moda (atacado/comissoes)")
    ap.add_argument("--endpoint", action="append", help="roda so este(s) endpoint(s)")
    ap.add_argument("--so-swagger", action="store_true", help="so le os swaggers de SWAGGER_URLS")
    args = ap.parse_args()

    api = Api()
    api.autenticar()

    swaggers = ler_swaggers(api) if SWAGGER_URLS else []
    resultados = []
    if not args.so_swagger:
        nomes = args.endpoint or list(ENDPOINTS)
        for nome in nomes:
            if nome not in ENDPOINTS:
                print(f"✘ endpoint desconhecido: {nome}. Opcoes: {', '.join(ENDPOINTS)}")
                continue
            resultados.append(sondar_endpoint(api, nome, ENDPOINTS[nome]))

    csv_path = salvar_inventario_csv(resultados)
    md_path = salvar_relatorio_md(resultados, swaggers)
    print(f"\nPronto.\n  Relatorio:  {md_path}\n  Inventario: {csv_path}\n  JSON bruto: {SAIDA}/")


if __name__ == "__main__":
    main()
