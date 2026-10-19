#!/usr/bin/env python3
"""
Faturas em Aberto (Atacado) - PDF no formato do relatório 0434, com TOTAL E BAIXAS
==================================================================================

Mesmo visual do 0434 do TOTVS (corretor > cliente > faturas, subtotais), mas com
o que Laura e Janaína pediram:
  - Transação preenchida MESMO APÓS BAIXA PARCIAL (no 0434 vira 0)
  - Cód. operação preenchido (no 0434 também vira 0)
  - Valor vendido original + total baixado + saldo em aberto
  - Histórico das baixas logo abaixo de cada fatura parcialmente paga

Uso:
  python relatorio_faturas_pdf.py                       # TODAS as compras do fechamento: vendido, baixado, saldo
  python relatorio_faturas_pdf.py --com-saldo           # só clientes que ainda devem, com TODAS as compras deles
  python relatorio_faturas_pdf.py --so-em-aberto        # igual ao 0434 (só faturas com saldo)
  python relatorio_faturas_pdf.py --sem-baixas          # sem o histórico de baixas
  python relatorio_faturas_pdf.py --venc-ini 2026-10-05 --venc-fim 2026-10-05
  python relatorio_faturas_pdf.py --corretor "LUCAS BALADA"   # só um corretor

Opcional no .env:  LOGO_PATH=C:\\caminho\\logo.png   (aparece no cabeçalho)

Saída: saida_api/faturas_em_aberto_<vencimento>_<data>.pdf
"""

from __future__ import annotations

import argparse
import os
import sys
from datetime import date, datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.units import mm
    from reportlab.platypus import SimpleDocTemplate, Spacer, Table, TableStyle, KeepTogether
except ImportError:
    sys.exit("Instale o reportlab:  pip install reportlab")

ap = argparse.ArgumentParser(description="PDF Faturas em Aberto (Atacado) com total e baixas")
ap.add_argument("--venc-ini")
ap.add_argument("--venc-fim")
ap.add_argument("--com-saldo", action="store_true",
                help="só clientes que ainda devem (mas lista TODAS as compras deles, inclusive as quitadas)")
ap.add_argument("--so-em-aberto", action="store_true",
                help="modo antigo do 0434: só as faturas com saldo (esconde as quitadas)")
ap.add_argument("--sem-baixas", action="store_true", help="não listar o histórico de baixas")
ap.add_argument("--corretor", action="append", help="filtrar corretor(es) pelo nome")
ARGS = ap.parse_args() if __name__ == "__main__" else ap.parse_args([])  # importado (web): padrões

import relatorio_faturas as RF  # noqa: E402  (usa --venc-ini/--venc-fim também)
S = RF.S

# ----------------------------------------------------------------------------
# Visual (cores do 0434)
# ----------------------------------------------------------------------------
AMARELO = colors.HexColor("#FFFF00")
AZUL = colors.HexColor("#9DC3E6")
CINZA = colors.HexColor("#808080")
CINZA_CLARO = colors.HexColor("#F2F2F2")
DESTAQUE = colors.HexColor("#FFF2CC")   # fatura parcialmente paga
FONTE, FONTE_B = "Courier", "Courier-Bold"
TAM = 7

COLS = ["Emp", "Vencimento", "Emissão", "Prz", "Nr. Fat", "Transação", "Oper",
        "Vl. vendido", "Vl. baixado", "Saldo", "Últ. baixa", "Situação"]
LARG = [10, 20, 20, 9, 17, 19, 10, 24, 24, 24, 20, 28]  # mm (soma ~ 225; A4 deitado útil ~ 267)
LARG = [w * mm * 1.18 for w in LARG]


def moeda(v) -> str:
    if v is None:
        return ""
    s = f"{v:,.2f}"
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def milhar(n) -> str:
    try:
        return f"{int(n):,}".replace(",", ".")
    except (TypeError, ValueError):
        return str(n or "")


def dt(d) -> str:
    if isinstance(d, (date, datetime)):
        return d.strftime("%d/%m/%Y")
    return str(d or "")


SIT_CURTA = {"Em aberto": "ABERTA", "Parcialmente paga": "PARCIAL", "Quitada": "QUITADA"}


# ----------------------------------------------------------------------------
# Montagem das linhas
# ----------------------------------------------------------------------------

def construir_story(faturas: list[dict], baixas: list[dict]):
    baixas_por_fat: dict[tuple, list] = {}
    for b in baixas:
        baixas_por_fat.setdefault((b["CD cliente"], b["Nr. Fat / NF"]), []).append(b)

    vendas = [f for f in faturas if f["Natureza"] == "Venda"]
    if ARGS.so_em_aberto:            # comportamento do 0434
        vendas = [f for f in vendas if f["Situação"] != "Quitada"]
    elif ARGS.com_saldo:             # cliente que ainda deve aparece com TODAS as compras
        saldo_cli: dict = {}
        for f in vendas:
            saldo_cli[f["CD cliente"]] = saldo_cli.get(f["CD cliente"], 0) + f["Saldo em aberto"]
        vendas = [f for f in vendas if saldo_cli.get(f["CD cliente"], 0) > 0.005]
    if ARGS.corretor:
        alvo = {c.strip().upper() for c in ARGS.corretor}
        vendas = [f for f in vendas if (f["Corretor"] or "").upper() in alvo]

    story = []
    tot = [0.0, 0.0, 0.0]
    corretores: dict[str, dict[tuple, list]] = {}
    for f in vendas:
        corretores.setdefault(f["Corretor"] or "(sem corretor)", {}).setdefault(
            (f["CD cliente"], f["Cliente"], f["CPF/CNPJ"], f.get("Status cliente")), []).append(f)

    for corretor, clientes in sorted(corretores.items()):
        linhas, estilo = [], []
        sub_cor = [0.0, 0.0, 0.0]

        def add(row, *cmds):
            linhas.append(row)
            i = len(linhas) - 1
            for c in cmds:
                estilo.append((c[0], (0, i), (-1, i), *c[1:]))

        add([f">: {corretor.lower()}"] + [""] * (len(COLS) - 1),
            ("SPAN",), ("BACKGROUND", AMARELO), ("FONTNAME", FONTE_B), ("ALIGN", "CENTER"),
            ("FONTSIZE", TAM + 1))

        manter_juntos = []
        for n_cli, ((cd, nome, doc, status), fats) in enumerate(sorted(clientes.items(), key=lambda kv: str(kv[0][1]))):
            ini_cli = 0 if n_cli == 0 else len(linhas)   # 1º cliente fica junto da faixa do corretor
            st = f" - Status: {str(status).upper()}" if status else ""
            add([f"CD: {cd} - {nome} - Cpfcnpj: {doc}{st}"] + [""] * (len(COLS) - 1),
                ("SPAN",), ("BACKGROUND", AZUL), ("FONTNAME", FONTE_B), ("ALIGN", "CENTER"))
            add(COLS, ("FONTNAME", FONTE_B), ("LINEBELOW", 0.3, colors.black))
            sub_cli = [0.0, 0.0, 0.0]
            for f in sorted(fats, key=lambda x: (str(x["Vencimento"]), str(x["Nr. Fat / NF"]))):
                parcial = f["Situação"] == "Parcialmente paga"
                cmds = [("BACKGROUND", DESTAQUE)] if parcial else []
                add([str(f.get("Empresa") or ""),
                     dt(f["Vencimento"]), dt(f["Emissão"]), str(f.get("Prz") or ""),
                     milhar(f["Nr. Fat / NF"]), milhar(f["Transação"]), str(f.get("Cód. Oper") or ""),
                     moeda(f["Valor vendido"]), moeda(f["Total baixado"]), moeda(f["Saldo em aberto"]),
                     dt(f["Última baixa"]), SIT_CURTA.get(f["Situação"], f["Situação"])], *cmds)
                if len(manter_juntos) == n_cli:          # 1ª fatura do cliente: prende ao cabeçalho
                    manter_juntos.append((ini_cli, len(linhas) - 1))
                sub_cli = [a + b for a, b in zip(sub_cli, (f["Valor vendido"], f["Total baixado"],
                                                           f["Saldo em aberto"]))]
                if not ARGS.sem_baixas and f["Total baixado"] > 0:
                    for b in sorted(baixas_por_fat.get((f["CD cliente"], f["Nr. Fat / NF"]), []),
                                    key=lambda x: (str(x["Data baixa"]), x["Parcela"] or 0)):
                        origem = f" (origem {b['Parcela origem']})" if b.get("Parcela origem") else ""
                        add(["", f"   baixa em {dt(b['Data baixa'])}", "", "", f"parc. {b['Parcela']}{origem}",
                             "", "", "", moeda(b["Valor baixado"]), "", "", b["Tipo de baixa"] or ""],
                            ("TEXTCOLOR", colors.HexColor("#555555")), ("FONTSIZE", TAM - 1),
                            ("SPAN_TXT",))
            add([f"SUB-TOTAL CLIENTE ({len(fats)} compra{'s' if len(fats) != 1 else ''}) ->", "", "", "", "", "", "",
                 moeda(sub_cli[0]), moeda(sub_cli[1]), moeda(sub_cli[2]), "", ""],
                ("FONTNAME", FONTE_B), ("FONTSIZE", TAM + 1), ("LINEABOVE", 0.3, colors.black))
            sub_cor = [a + b for a, b in zip(sub_cor, sub_cli)]

        add([f"SUB-TOTAL {corretor} ->", "", "", "", "", "", "",
             moeda(sub_cor[0]), moeda(sub_cor[1]), moeda(sub_cor[2]), "", ""],
            ("BACKGROUND", CINZA), ("TEXTCOLOR", colors.white), ("FONTNAME", FONTE_B), ("FONTSIZE", TAM + 1))
        tot = [a + b for a, b in zip(tot, sub_cor)]

        # converte comandos simplificados em TableStyle
        ts = [("FONTNAME", (0, 0), (-1, -1), FONTE), ("FONTSIZE", (0, 0), (-1, -1), TAM),
              ("ALIGN", (3, 0), (-1, -1), "RIGHT"), ("ALIGN", (-1, 0), (-1, -1), "CENTER"),
              ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
              ("LEFTPADDING", (0, 0), (-1, -1), 2), ("RIGHTPADDING", (0, 0), (-1, -1), 2)]
        for cmd in estilo:
            nome, ini, fim, *rest = cmd
            if nome == "SPAN":
                ts.append(("SPAN", ini, fim))
            elif nome == "SPAN_TXT":  # linha de baixa: texto da data ocupa colunas 1-3
                ts.append(("SPAN", (1, ini[1]), (3, ini[1])))
                ts.append(("ALIGN", (1, ini[1]), (4, ini[1]), "LEFT"))
            else:
                ts.append((nome, ini, fim, *rest))
        # rótulos de subtotal à esquerda ocupando as 7 primeiras colunas
        for i, row in enumerate(linhas):
            if str(row[0]).startswith("SUB-TOTAL"):
                ts.append(("SPAN", (0, i), (6, i)))
                ts.append(("ALIGN", (0, i), (6, i), "LEFT"))
        for a, b in manter_juntos:
            ts.append(("NOSPLIT", (0, a), (-1, b)))
        for i, row in enumerate(linhas):          # subtotal nunca sozinho no topo da página
            if str(row[0]).startswith("SUB-TOTAL") and i > 0:
                ts.append(("NOSPLIT", (0, i - 1), (-1, i)))
        t = Table(linhas, colWidths=LARG, repeatRows=0)
        t.setStyle(TableStyle(ts))
        story.append(t)
        story.append(Spacer(1, 3 * mm))

    # total geral
    t = Table([["TOTAL GERAL ->", "", "", "", "", "", "", moeda(tot[0]), moeda(tot[1]), moeda(tot[2]), "", ""]],
              colWidths=LARG)
    t.setStyle(TableStyle([("SPAN", (0, 0), (6, 0)), ("FONTNAME", (0, 0), (-1, -1), FONTE_B),
                           ("FONTSIZE", (0, 0), (-1, -1), TAM + 2), ("ALIGN", (7, 0), (9, 0), "RIGHT"),
                           ("BACKGROUND", (0, 0), (-1, -1), colors.black), ("TEXTCOLOR", (0, 0), (-1, -1), colors.white)]))
    story.append(KeepTogether([t]))

    resumo = {"faturas": len(vendas), "vendido": tot[0], "baixado": tot[1], "saldo": tot[2],
              "parciais": sum(1 for f in vendas if f["Situação"] == "Parcialmente paga"),
              "corretores": len(corretores),
              "clientes": len({f["CD cliente"] for f in vendas})}
    return story, resumo


# ----------------------------------------------------------------------------
# Cabeçalho / rodapé (estilo 0434)
# ----------------------------------------------------------------------------

def cabecalho_rodape(resumo: dict, gerado_em: datetime):
    titulo = ("FATURAS EM ABERTO (ATACADO)" if ARGS.so_em_aberto
              else "FATURAS DO ATACADO - TOTAL VENDIDO E BAIXAS")
    filtros = [
        f"Empresa = {S.BRANCHES} /",
        f"Vencimento = [>= {RF.date.fromisoformat(S.VENC_INI):%d/%m/%Y} & <= "
        f"{RF.date.fromisoformat(S.VENC_FIM):%d/%m/%Y}] /",
        f"Fatura em = [{'Em aberto' if ARGS.so_em_aberto else 'Todas (abertas e baixadas)'}] /",
        f"Clientes = [{'Com saldo' if ARGS.com_saldo else 'Todos'}] / Situação = [Normal] /",
        "Operação = [800 Venda atacado, 802 Devolução]"
        + (f" / Corretor = {ARGS.corretor}" if ARGS.corretor else ""),
    ]
    logo = os.getenv("LOGO_PATH")

    def desenhar(canv, doc):
        w, h = doc.pagesize
        canv.saveState()
        if logo and Path(logo).exists():
            try:
                canv.drawImage(logo, 12 * mm, h - 30 * mm, width=28 * mm, height=18 * mm,
                               preserveAspectRatio=True, mask="auto")
            except Exception:  # noqa: BLE001
                pass
        canv.setFont(FONTE_B, 13)
        canv.drawCentredString(w / 2, h - 12 * mm, titulo)
        canv.setFont(FONTE, 7.5)
        y = h - 17 * mm
        canv.drawString(w / 2 - 95 * mm, y, "Filtro: " + filtros[0])
        for f in filtros[1:]:
            y -= 3.3 * mm
            canv.drawString(w / 2 - 95 * mm + 12 * mm, y, f)
        # quadro-resumo no canto direito (só na 1ª página)
        if doc.page == 1:
            x = w - 70 * mm
            canv.setFont(FONTE_B, 7.5)
            canv.drawString(x, h - 14 * mm, "RESUMO")
            canv.setFont(FONTE, 7.5)
            itens = [("Clientes / compras", f'{resumo["clientes"]} / {resumo["faturas"]}'),
                     ("Parcialmente pagas", str(resumo["parciais"])),
                     ("Vendido", moeda(resumo["vendido"])), ("Baixado", moeda(resumo["baixado"])),
                     ("Saldo em aberto", moeda(resumo["saldo"]))]
            for i, (k, v) in enumerate(itens):
                yy = h - (17.5 + 3.3 * i) * mm
                canv.drawString(x, yy, k)
                canv.drawRightString(w - 12 * mm, yy, v)
        # legenda + rodapé
        canv.setFont(FONTE, 6.5)
        canv.setFillColor(DESTAQUE)
        canv.rect(12 * mm, 7.5 * mm, 4 * mm, 2.6 * mm, fill=1, stroke=0)
        canv.setFillColor(colors.black)
        canv.drawString(17 * mm, 8 * mm, "= fatura com baixa parcial (no 0434 aparece com Transação 0)")
        canv.line(12 * mm, 11.5 * mm, w - 12 * mm, 11.5 * mm)
        canv.drawRightString(w - 12 * mm, 8 * mm,
                             f"Gerado via API TOTVS Moda / Dat:{gerado_em:%d/%m/%y %H:%M} / Página: {doc.page}")
        canv.restoreState()
    return desenhar


def nome_arquivo(agora: datetime) -> str:
    venc = S.VENC_INI.replace("-", "") + ("" if S.VENC_INI == S.VENC_FIM else "-" + S.VENC_FIM.replace("-", ""))
    nome = "faturas_em_aberto" if ARGS.so_em_aberto else "faturas_total_e_baixas"
    return f"{nome}_{venc}_{agora:%Y-%m-%d_%H%M}.pdf"


def gerar(destino, agora: datetime) -> dict:
    """Busca na API e grava o PDF em `destino` (caminho ou arquivo em memória). Devolve o resumo."""
    linhas = RF.coletar()
    faturas, baixas = RF.agrupar(linhas)
    story, resumo = construir_story(faturas, baixas)
    doc = SimpleDocTemplate(destino, pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm,
                            topMargin=37 * mm, bottomMargin=15 * mm,
                            title="Faturas em Aberto (Atacado) - Total e Baixas", author="Bebêtenkitê")
    pintar = cabecalho_rodape(resumo, agora)
    doc.build(story, onFirstPage=pintar, onLaterPages=pintar)
    return resumo


def main() -> None:
    agora = datetime.now()
    S.SAIDA.mkdir(exist_ok=True)
    arq = S.SAIDA / nome_arquivo(agora)
    resumo = gerar(str(arq), agora)
    print(f"\n✔ {resumo['faturas']} faturas | {resumo['corretores']} corretores | "
          f"vendido R$ {moeda(resumo['vendido'])} | baixado R$ {moeda(resumo['baixado'])} | "
          f"saldo R$ {moeda(resumo['saldo'])} | {resumo['parciais']} parcialmente pagas")
    print(f"  PDF: {arq}")


if __name__ == "__main__":
    main()
