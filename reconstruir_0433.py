#!/usr/bin/env python3
"""
Reconstroi o relatorio 0433 (Controle Atacado) a partir da API TOTVS Moda
e compara com a planilha atual.
==========================================================================

Hipoteses que este script testa (vindas da sonda):
  H1. Nr. Fat  = accounts-receivable.receivableCode  (= numero da NF)
  H2. Transacao = fiscal.invoices.transactionCode, ligada ao titulo por
      receivable.invoice[].invoiceCode + invoiceSequence
  H3. Corretor = campo "homePage" do cadastro do cliente (PF/PJ)
      (o Power Query so passa para maiusculo)
  H4. Atacado  = NF com operacao 800 (VENDA ATACADO) / 802 (DEVOLUCAO)
  H5. Baixa parcial cria parcela nova (installmentCode >= 200) com
      originInstallment apontando para a parcela original
  H6. Tp. Baixa = dischargeType (0 = em aberto)

Uso:
  python reconstruir_0433.py                                  # so API
  python reconstruir_0433.py --planilha "00. CONTROLE VENDAS ATACADO.xlsm"
  python reconstruir_0433.py --planilha BEBETENKITE_6_MALENCAR_0433.CSV

Saidas (pasta saida_api/):
  controle_atacado_api.csv   -> mesmas colunas do 0433 + colunas extras
  titulos_nao_classificados.csv -> titulos do vencimento que nao ligaram a NF de atacado
  comparacao.md              -> o que bate, o que falta, transacoes recuperadas

Usa a mesma configuracao (.env) do sonda_api_totvs.py.
Extra opcional no .env:
  OPERACOES_ATACADO=800,802
  MAX_PAGINAS_CARGA=50
"""

from __future__ import annotations

import argparse
import csv
import os
import sys
from collections import Counter, defaultdict
from datetime import date
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import sonda_api_totvs as S  # noqa: E402  (reaproveita Api, .env, datas)

OPERACOES_ATACADO = {int(x) for x in os.getenv("OPERACOES_ATACADO", "800,802").split(",") if x.strip()}
MAX_PAGINAS_CARGA = int(os.getenv("MAX_PAGINAS_CARGA", "500"))  # 50 cortava períodos longos (6 meses ~ 62 páginas)

PATH_AR = "/api/totvsmoda/accounts-receivable/v2/documents/search"
PATH_NF = "/api/totvsmoda/fiscal/v2/invoices/search"
PATH_PJ = "/api/totvsmoda/person/v2/legal-entities/search"
PATH_PF = "/api/totvsmoda/person/v2/individuals/search"


# ----------------------------------------------------------------------------
# Coleta paginada
# ----------------------------------------------------------------------------

PROGRESSO = None  # opcional: função (rotulo, pagina, total_paginas) chamada a cada página (usada pelo app.py)


def buscar_tudo(api: S.Api, path: str, filtro: dict, expand: str | None, rotulo: str) -> list[dict]:
    itens: list[dict] = []
    for pagina in range(1, MAX_PAGINAS_CARGA + 1):
        body = {"filter": filtro, "page": pagina, "pageSize": S.PAGE_SIZE}
        if expand:
            body["expand"] = expand
        r = api.chamar("POST", path, body)
        if not r.ok:
            sys.exit(f"✘ {rotulo}: HTTP {r.status_code} {r.text[:500]}")
        dados = r.json()
        itens.extend(dados.get("items") or [])
        print(f"   {rotulo}: pagina {pagina}/{dados.get('totalPages', '?')} -> {len(itens)} itens")
        if PROGRESSO:
            PROGRESSO(rotulo, pagina, dados.get("totalPages") or pagina)
        if not dados.get("hasNext"):
            break
    else:
        print(f"   ⚠ {rotulo}: parou em MAX_PAGINAS_CARGA={MAX_PAGINAS_CARGA} (pode haver mais)")
    return itens


def buscar_pessoas(api: S.Api, codigos: set[int]) -> dict[int, dict]:
    pessoas: dict[int, dict] = {}
    lista = sorted(codigos)
    for i in range(0, len(lista), 50):
        lote = lista[i:i + 50]
        for path, tipo in ((PATH_PJ, "PJ"), (PATH_PF, "PF")):
            for p in buscar_tudo(api, path, {"personCodeList": lote}, None, f"pessoas {tipo} {i + 1}-{i + len(lote)}"):
                cod = p.get("code")
                if cod is not None:
                    pessoas[cod] = {
                        "nome": p.get("name"),
                        "doc": p.get("cnpj") or p.get("cpf"),
                        "homePage": (p.get("homePage") or "").strip(),
                        "representatives": p.get("representatives"),
                        "tipo": tipo,
                    }
    return pessoas


# ----------------------------------------------------------------------------
# Montagem
# ----------------------------------------------------------------------------

def d_br(iso: str | None) -> str:
    if not iso:
        return ""
    y, m, d = iso[:10].split("-")
    return f"{d}/{m}/{y}"


def prazo(emissao: str | None, venc: str | None) -> str:
    try:
        return str((date.fromisoformat(venc[:10]) - date.fromisoformat(emissao[:10])).days)
    except Exception:  # noqa: BLE001
        return ""


COLUNAS_0433 = ["Empresa", "Corretor / Representante", "CD cliente", "Cliente", "Cpfcnpj",
                "Vencimento", "Emissao", "Prz", "Nr. Fat", "Dt. Baixa", "Transação", "Valor", "Tp. Baixa"]
NATUREZA = {1: "Venda", 20: "Crédito devolução (Credev)"}  # documentType (swagger)
DOC_TIPO = {1: "Fatura", 2: "Cheque", 3: "Dinheiro", 4: "Cartão de crédito", 5: "Cartão de débito",
            6: "Nota de débito", 20: "CREDEV"}
STATUS_TITULO = {1: "Normal", 2: "Devolvido", 3: "Cancelado", 4: "Quebrada"}
# dischargeType: 0-4 pela descricao do swagger; 10 e 12 pela ordem do enum (confirmar)
TIPO_BAIXA = {0: "Em aberto", 1: "Via recebimento", 2: "Baixa para depósito", 3: "Agrupamento de fatura",
              4: "Devolução", 9: "Renegociação?", 10: "Via faturamento?", 12: "Conta bancária?"}
STATUS_EXCLUIR = {int(x) for x in os.getenv("STATUS_EXCLUIR", "3").split(",") if x.strip()}  # 3 = Cancelado

COLUNAS_EXTRA = ["Natureza", "Tipo baixa (nome)", "Parcela", "Parcela origem", "Valor pago", "Status titulo", "Tipo documento",
                 "Portador", "Operacao NF", "Valor total NF", "Transacao recuperada?",
                 "Corretor no cadastro (representatives)"]


def montar(titulos: list[dict], notas: list[dict], pessoas: dict[int, dict]):
    # indice de NF: (invoiceCode, invoiceSequence) e so invoiceCode
    nf_por_seq: dict[tuple, dict] = {}
    nf_por_cod: dict[int, list[dict]] = defaultdict(list)
    for nf in notas:
        nf_por_seq[(nf.get("invoiceCode"), nf.get("invoiceSequence"))] = nf
        nf_por_cod[nf.get("invoiceCode")].append(nf)

    # transacao por (receivableCode, parcela) para herdar na parcela residual
    linhas, nao_class = [], []
    for t in titulos:
        nf, via = None, ""
        for inv in t.get("invoice") or []:
            nf = nf_por_seq.get((inv.get("invoiceCode"), inv.get("invoiceSequence")))
            if nf:
                via = "invoice[]"
                break
            cands = nf_por_cod.get(inv.get("invoiceCode")) or []
            if len(cands) == 1:
                nf, via = cands[0], "invoiceCode"
                break
        if not nf:  # fallback: Nr. Fat = numero da NF
            cands = nf_por_cod.get(t.get("receivableCode")) or []
            if len(cands) == 1:
                nf, via = cands[0], "receivableCode"
        t["_nf"], t["_via"] = nf, via

    # agrupa por fatura para herdar a NF da parcela original (H5)
    por_fatura: dict[tuple, list[dict]] = defaultdict(list)
    for t in titulos:
        por_fatura[(t.get("customerCode"), t.get("receivableCode"))].append(t)
    for grupo in por_fatura.values():
        com_nf = [t for t in grupo if t["_nf"]]
        for t in grupo:
            if not t["_nf"] and com_nf:
                t["_nf"], t["_via"] = com_nf[0]["_nf"], "herdada da parcela original"

    for t in titulos:
        nf = t["_nf"]
        if not nf or nf.get("operationCode") not in OPERACOES_ATACADO:
            nao_class.append(t)
            continue
        pes = pessoas.get(t.get("customerCode"), {})
        tipo_doc = t.get("documentType")
        linhas.append({
            "Natureza": NATUREZA.get(tipo_doc, f"Outro ({tipo_doc})"),
            "Empresa": t.get("branchCode"),
            "Corretor / Representante": pes.get("homePage", "").upper(),
            "CD cliente": t.get("customerCode"),
            "Cliente": pes.get("nome") or nf.get("personName"),
            "Cpfcnpj": t.get("customerCpfCnpj"),
            "Vencimento": d_br(t.get("expiredDate")),
            "Emissao": d_br(t.get("issueDate")),
            "Prz": prazo(t.get("issueDate"), t.get("expiredDate")),
            "Nr. Fat": t.get("receivableCode"),
            "Dt. Baixa": d_br(t.get("settlementDate")),
            "Transação": nf.get("transactionCode"),
            "Valor": t.get("installmentValue"),
            "Tp. Baixa": t.get("dischargeType") or 0,
            "Parcela": t.get("installmentCode"),
            "Parcela origem": t.get("originInstallment") or "",
            "Valor pago": t.get("paidValue"),
            "Status titulo": STATUS_TITULO.get(t.get("status"), t.get("status")),
            "Tipo documento": DOC_TIPO.get(t.get("documentType"), t.get("documentType")),
            "Tipo baixa (nome)": TIPO_BAIXA.get(t.get("dischargeType") or 0, t.get("dischargeType")),
            "Portador": t.get("bearerName"),
            "Operacao NF": nf.get("operatioName") or nf.get("operationCode"),
            "Valor total NF": nf.get("totalValue"),
            "Transacao recuperada?": t["_via"],
            "Corretor no cadastro (representatives)": bool(pes.get("representatives")),
        })
    linhas.sort(key=lambda r: (r["Corretor / Representante"], str(r["Cliente"]), str(r["Nr. Fat"])))
    return linhas, nao_class


# ----------------------------------------------------------------------------
# Planilha atual
# ----------------------------------------------------------------------------

def ler_planilha(caminho: Path) -> list[dict]:
    cols = ["Nome da Origem", "Empresa", "Corretor / Representante", "CD cliente", "Cliente", "Cpfcnpj",
            "Vencimento", "Emissao", "Prz", "Nr. Fat", "Dt. Baixa", "Transação", "Valor", "Tp. Baixa"]
    if caminho.suffix.lower() in (".xlsm", ".xlsx"):
        try:
            import openpyxl
        except ImportError:
            sys.exit("Para ler .xlsm: pip install openpyxl")
        ws = openpyxl.load_workbook(caminho, read_only=True, data_only=True)["DADOS"]
        linhas = [dict(zip(cols, r)) for r in ws.iter_rows(min_row=2, values_only=True)]
    else:  # CSV 0433 bruto (13 colunas, ; , cp1252, coluna ">" = corretor)
        with caminho.open(encoding="cp1252", newline="") as f:
            rd = csv.reader(f, delimiter=";")
            cab = next(rd)
            linhas = []
            for r in rd:
                d = dict(zip(cols[1:], r))
                linhas.append(d)
        _ = cab
    out = []
    for r in linhas:
        if not r.get("Nr. Fat"):
            continue
        try:
            valor = round(float(str(r.get("Valor")).replace(".", "").replace(",", ".")
                                if isinstance(r.get("Valor"), str) and "," in str(r.get("Valor"))
                                else r.get("Valor")), 2)
        except (TypeError, ValueError):
            valor = None
        out.append({**r, "_valor": valor, "_fat": str(r["Nr. Fat"]).strip(),
                    "_cli": str(r.get("CD cliente")).strip()})
    return out


def comparar(api_linhas: list[dict], plan: list[dict]) -> list[str]:
    L = ["\n## Comparação com a planilha\n"]
    chave = lambda cli, fat, v: (str(cli), str(fat), round(float(v), 2) if v is not None else None)  # noqa: E731
    api_idx = defaultdict(list)
    for r in api_linhas:
        api_idx[chave(r["CD cliente"], r["Nr. Fat"], r["Valor"])].append(r)
    plan_idx = defaultdict(list)
    for r in plan:
        plan_idx[chave(r["_cli"], r["_fat"], r["_valor"])].append(r)

    so_plan = [k for k in plan_idx if k not in api_idx]
    so_api = [k for k in api_idx if k not in plan_idx]
    ambos = [k for k in plan_idx if k in api_idx]

    # titulos desmembrados depois da exportacao: soma das parcelas da API = valor da planilha
    desmembrados = []
    api_sobra = defaultdict(list)
    for k in so_api:
        api_sobra[(k[0], k[1])].append(k)
    for k in list(so_plan):
        partes = api_sobra.get((k[0], k[1]), [])
        soma = round(sum(p[2] for p in partes), 2)
        if partes and k[2] is not None and abs(soma - k[2]) < 0.02:
            desmembrados.append((k, partes))
            so_plan.remove(k)
            for p in partes:
                so_api.remove(p)

    L.append(f"- Linhas na planilha: **{len(plan)}** | na API (atacado): **{len(api_linhas)}**")
    L.append(f"- Batem (cliente + Nr. Fat + valor): **{len(ambos)}**")
    L.append(f"- Desmembradas depois da exportação (soma das parcelas = valor da planilha): **{len(desmembrados)}**")
    L.append(f"- Sobram só na planilha: **{len(so_plan)}** | só na API: **{len(so_api)}**\n")
    if desmembrados:
        L.append("### Títulos desmembrados por baixa parcial após a exportação\n")
        for k, partes in desmembrados:
            pp = " + ".join(f"{p[2]:.2f} (parc. {api_idx[p][0]['Parcela']})" for p in partes)
            L.append(f"- cliente {k[0]} fat {k[1]}: R$ {k[2]:.2f} = {pp}")
        L.append("")

    div_corretor, div_trans, recuperadas, mudou_baixa = [], [], [], []
    for k in ambos:
        p, a = plan_idx[k][0], api_idx[k][0]
        if str(p.get("Corretor / Representante") or "").strip().upper() != a["Corretor / Representante"]:
            div_corretor.append((k, p.get("Corretor / Representante"), a["Corretor / Representante"]))
        tp, ta = str(p.get("Transação") or "").strip(), str(a["Transação"] or "")
        if tp in ("0", "") and ta not in ("", "0", "None"):
            recuperadas.append((k, ta, a["Transacao recuperada?"]))
        elif tp != ta:
            div_trans.append((k, tp, ta))
        if str(p.get("Tp. Baixa") or "0") != str(a["Tp. Baixa"]):
            mudou_baixa.append((k, p.get("Tp. Baixa"), a["Tp. Baixa"], a["Dt. Baixa"]))

    L.append(f"### Corretor (H3 - homePage): {len(ambos) - len(div_corretor)}/{len(ambos)} batem\n")
    for k, p, a in div_corretor[:30]:
        L.append(f"- cliente {k[0]} fat {k[1]} R$ {k[2]}: planilha `{p}` x API `{a}`")
    L.append(f"\n### Transação (H2): {len(ambos) - len(div_trans) - len(recuperadas)} iguais, "
             f"**{len(recuperadas)} recuperadas** (eram 0 na planilha), {len(div_trans)} diferentes\n")
    for k, ta, via in recuperadas:
        L.append(f"- ✅ cliente {k[0]} fat {k[1]} R$ {k[2]}: 0 → **{ta}** (via {via})")
    for k, tp, ta in div_trans[:30]:
        L.append(f"- ⚠ cliente {k[0]} fat {k[1]} R$ {k[2]}: planilha {tp} x API {ta}")
    L.append(f"\n### Tp. Baixa mudou desde a exportação do CSV: {len(mudou_baixa)}\n")
    L.append("(Esperado: títulos que estavam em aberto e foram pagos depois.)\n")
    for k, p, a, dt in mudou_baixa[:40]:
        L.append(f"- cliente {k[0]} fat {k[1]} R$ {k[2]}: {p} → {a} (baixa {dt})")
    if so_plan:
        L.append("\n### Só na planilha (investigar)\n")
        for k in so_plan[:40]:
            p = plan_idx[k][0]
            L.append(f"- {p.get('Corretor / Representante')} | cliente {k[0]} | fat {k[1]} | R$ {k[2]} | "
                     f"transação {p.get('Transação')} | tp {p.get('Tp. Baixa')}")
    if so_api:
        L.append("\n### Só na API (investigar: parcela nova? devolução? lançado depois?)\n")
        for k in so_api[:40]:
            a = api_idx[k][0]
            L.append(f"- {a['Corretor / Representante']} | cliente {k[0]} | fat {k[1]} | R$ {k[2]} | "
                     f"**{a['Natureza']}** | status {a['Status titulo']} | tp baixa {a['Tp. Baixa']} | "
                     f"emissão {a['Emissao']} | parcela {a['Parcela']} (origem {a['Parcela origem']})")
        st = Counter(api_idx[k][0]["Status titulo"] for k in so_api if api_idx[k][0]["Natureza"] == "Venda")
        L.append(f"\nStatus dos títulos de venda que só existem na API: {dict(st)} "
                 f"(se todos forem 3, o relatório 0433 provavelmente filtra esse status)")
    return L


def ler_regra(caminho: Path) -> set[str]:
    """Nomes de corretor da aba REGRA (so para .xlsm/.xlsx)."""
    if caminho.suffix.lower() not in (".xlsm", ".xlsx"):
        return set()
    import openpyxl
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    if "REGRA" not in wb.sheetnames:
        return set()
    return {str(r[1]).strip().upper() for r in wb["REGRA"].iter_rows(min_row=3, values_only=True)
            if len(r) > 1 and r[1]}


# ----------------------------------------------------------------------------

def salvar_csv(arq: Path, linhas: list[dict], colunas: list[str]) -> None:
    with arq.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=colunas, delimiter=";", extrasaction="ignore")
        w.writeheader()
        for r in linhas:
            w.writerow({k: (str(v).replace(".", ",") if isinstance(v, float) else v) for k, v in r.items()})


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--planilha", help="xlsm do Controle Atacado (aba DADOS) ou CSV 0433 para comparar")
    args = ap.parse_args()

    api = S.Api()
    api.autenticar()
    S.SAIDA.mkdir(exist_ok=True)

    print("\n▶ Títulos a receber do vencimento")
    titulos = buscar_tudo(api, PATH_AR, {
        "branchCodeList": S.BRANCHES,
        "startExpiredDate": S.data_iso(S.VENC_INI),
        "endExpiredDate": S.data_iso(S.VENC_FIM, fim=True),
    }, "invoice,calculateValue", "títulos")

    # periodo de NF: da emissao mais antiga dos titulos ate a mais nova
    emissoes = sorted(t["issueDate"][:10] for t in titulos if t.get("issueDate"))
    nf_ini = emissoes[0] if emissoes else S.EMISSAO_INI
    nf_fim = emissoes[-1] if emissoes else S.EMISSAO_FIM
    print(f"\n▶ Notas fiscais emitidas de {nf_ini} a {nf_fim}")
    notas = buscar_tudo(api, PATH_NF, {
        "branchCodeList": S.BRANCHES,
        "operationCodeList": sorted(OPERACOES_ATACADO),   # filtro existe no swagger
        "startIssueDate": S.data_iso(nf_ini),
        "endIssueDate": S.data_iso(nf_fim, fim=True),
    }, None, "notas")
    notas_canc = [n for n in notas if n.get("invoiceStatus") in ("Canceled", "Deleted", "Denied")]
    notas = [n for n in notas if n not in notas_canc]
    ops = Counter((n.get("operationCode"), n.get("operatioName")) for n in notas)

    # pessoas so dos clientes com NF de atacado
    clientes_atacado = {n.get("personCode") for n in notas if n.get("operationCode") in OPERACOES_ATACADO}
    print(f"\n▶ Cadastro de {len(clientes_atacado)} clientes do atacado")
    pessoas = buscar_pessoas(api, clientes_atacado)

    cancelados = [t for t in titulos if t.get("status") in STATUS_EXCLUIR]
    titulos = [t for t in titulos if t.get("status") not in STATUS_EXCLUIR]
    linhas, nao_class = montar(titulos, notas, pessoas)
    salvar_csv(S.SAIDA / "controle_atacado_api.csv", linhas, COLUNAS_0433 + COLUNAS_EXTRA)
    salvar_csv(S.SAIDA / "titulos_nao_classificados.csv",
               [{"Nr. Fat": t.get("receivableCode"), "Parcela": t.get("installmentCode"),
                 "CD cliente": t.get("customerCode"), "Valor": t.get("installmentValue"),
                 "Portador": t.get("bearerName"), "Tipo documento": t.get("documentType"),
                 "Operacao NF": (t["_nf"] or {}).get("operatioName", "sem NF ligada")} for t in nao_class],
               ["Nr. Fat", "Parcela", "CD cliente", "Valor", "Portador", "Tipo documento", "Operacao NF"])

    # ---------------- relatorio ----------------
    L = ["# Reconstrução do relatório 0433 pela API",
         f"\nVencimento {S.VENC_INI} a {S.VENC_FIM} | empresa {S.BRANCHES} | operações atacado {sorted(OPERACOES_ATACADO)}\n",
         "## Volumes\n",
         f"- Títulos com esse vencimento: **{len(titulos)}**",
         f"- Notas fiscais no período: **{len(notas)}**",
         f"- Títulos de atacado montados: **{len(linhas)}** "
         f"(R$ {sum(float(r['Valor'] or 0) for r in linhas):,.2f})",
         f"- Títulos não classificados como atacado: **{len(nao_class)}** (ver titulos_nao_classificados.csv)",
         f"- Títulos **cancelados** (status {sorted(STATUS_EXCLUIR)}) excluídos: **{len(cancelados)}** "
         f"(R$ {sum(float(t.get('installmentValue') or 0) for t in cancelados):,.2f})",
         f"- Notas canceladas/excluídas/denegadas ignoradas: **{len(notas_canc)}**\n",
         "### Operações fiscais encontradas\n"]
    L += [f"- {cod} {nome}: {q} notas" for (cod, nome), q in ops.most_common()]
    L.append("\n### Títulos não classificados, por portador\n")
    L += [f"- {b}: {q}" for b, q in Counter(t.get("bearerName") for t in nao_class).most_common(15)]

    L.append("\n## Hipóteses\n")
    sem_corretor = [r for r in linhas if not r["Corretor / Representante"]]
    L.append(f"- **H3 corretor = homePage**: {len(linhas) - len(sem_corretor)}/{len(linhas)} títulos com corretor; "
             f"{len(sem_corretor)} sem (clientes: {sorted({r['CD cliente'] for r in sem_corretor})[:20]})")
    L.append(f"- Corretores encontrados: {sorted({r['Corretor / Representante'] for r in linhas if r['Corretor / Representante']})}")
    via = Counter(r["Transacao recuperada?"] for r in linhas)
    L.append(f"- **H2 transação**: ligação título→NF por {dict(via)}")
    resid = [r for r in linhas if (r["Parcela"] or 0) >= 200]
    L.append(f"- **H5 parcelas residuais (>=200)**: {len(resid)}; com parcela origem preenchida: "
             f"{sum(1 for r in resid if r['Parcela origem'])}")
    L.append(f"- **Tipo de baixa**: {dict(Counter(r['Tipo baixa (nome)'] for r in linhas))}")
    L.append(f"- Status do título: {dict(Counter(r['Status titulo'] for r in linhas))} | "
             f"Tipo documento: {dict(Counter(r['Tipo documento'] for r in linhas))}")
    L.append(f"- Clientes com `representatives` preenchido no cadastro: "
             f"{sum(1 for p in pessoas.values() if p['representatives'])}/{len(pessoas)}")

    # vendas e creditos de devolucao separados (Credev nao e venda)
    por_corretor = defaultdict(lambda: [0.0, 0.0, 0.0, 0.0])
    for r in linhas:
        v = float(r["Valor"] or 0)
        c = por_corretor[r["Corretor / Representante"] or "(sem corretor)"]
        if r["Natureza"] != "Venda":
            c[3] += v
            continue
        c[0] += v
        if str(r["Tp. Baixa"]) == "0":
            c[2] += v
        else:
            c[1] += v
    L.append("\n## Prévia da aba TD (só vendas; Credev à parte)\n")
    L.append("| Corretor | Vendido | Pago | Em aberto | Créditos devolução |\n|---|---:|---:|---:|---:|")
    tot = [0.0] * 4
    for nome, vals in sorted(por_corretor.items()):
        tot = [a + b for a, b in zip(tot, vals)]
        L.append(f"| {nome} | " + " | ".join(f"{x:,.2f}" for x in vals) + " |")
    L.append("| **Total** | " + " | ".join(f"**{x:,.2f}**" for x in tot) + " |")

    if args.planilha:
        plan = ler_planilha(Path(args.planilha))
        regra = ler_regra(Path(args.planilha))
        if regra:
            fora = sorted({r["Corretor / Representante"] for r in linhas} - regra - {""})
            L.append(f"\n## Corretores sem regra de comissão (aba REGRA)\n")
            L.append(f"- {fora if fora else 'nenhum ✅'}"
                     + (" → comissão sairia R$ 0 na planilha atual" if fora else ""))
        L += comparar(linhas, plan)

    (S.SAIDA / "comparacao.md").write_text("\n".join(L), encoding="utf-8")
    print(f"\nPronto: {S.SAIDA / 'comparacao.md'}\n       {S.SAIDA / 'controle_atacado_api.csv'}")


if __name__ == "__main__":
    main()
