"""Gera as agendas para assinar (arquivos .ics) a partir da planilha pública do Calendário Olímpico.

Roda no GitHub Actions (veja .github/workflows/agendas.yml). Lê a planilha pelo mesmo
serviço público que o site usa e grava um arquivo por segmento em agendas/.
Só entram datas públicas: prazos internos, reuniões e "Somente equipe" ficam de fora.

Teste local com dados de exemplo:  python scripts/gerar_agendas.py --exemplo dados.json
"""
import hashlib
import json
import re
import sys
import unicodedata
import urllib.parse
import urllib.request
from datetime import date, timedelta
from pathlib import Path

SHEET_ID = "1wF10E2_omNchLjgwAkMNBRixaEMf1wp17DaxzZsgKag"
SITE = "https://ctomackenziebrasilia.github.io/"
SAIDA = Path(__file__).resolve().parent.parent / "agendas"

AGENDAS = {
    "efai": ("Olimpíadas CTO · Anos Iniciais (1º ao 5º ano)", "EFAI"),
    "efaf": ("Olimpíadas CTO · Anos Finais (6º ao 9º ano)", "EFAF"),
    "em": ("Olimpíadas CTO · Ensino Médio", "EM"),
}
ETAPAS = [  # (nome, coluna início, coluna fim, interna)
    ("Inscrição", "inscricao (inicio)", "inscricao (fim)", False),
    ("Divulgação aos alunos", "divulgacao aos alunos", None, False),
    ("Prazo interno CTO", "prazo interno cto", None, True),
    ("1ª Fase", "1ª fase (inicio)", "1ª fase (fim)", False),
    ("2ª Fase", "2ª fase (inicio)", "2ª fase (fim)", False),
    ("Fase Final", "fase final (inicio)", "fase final (fim)", False),
    ("Resultado", "resultado", None, False),
    ("Premiação", "premiacao", None, False),
]
TIPOS_INTERNOS = {"prazo interno", "prazo interno cto", "reuniao"}
MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
         "agosto", "setembro", "outubro", "novembro", "dezembro"]
JANELA_CURTA = 14  # períodos de até 14 dias viram um evento só; mais longos, início e fim


def norm(v):
    return "" if v is None else str(v).strip()


def key(v):
    s = unicodedata.normalize("NFD", norm(v).lower())
    return "".join(c for c in s if unicodedata.category(c) != "Mn")


def para_data(v):
    v = norm(v)
    if not v:
        return None
    m = re.match(r"^Date\((\d+),(\d+),(\d+)", v)
    if m:
        return date(int(m[1]), int(m[2]) + 1, int(m[3]))
    m = re.match(r"^(\d{4})-(\d{2})-(\d{2})", v)
    if m:
        return date(int(m[1]), int(m[2]), int(m[3]))
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", v)
    if m:
        return date(int(m[3]), int(m[2]), int(m[1]))
    return None


def ler_aba(nome):
    url = (f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?headers=1&tqx=out:json&sheet="
           + urllib.parse.quote(nome))
    with urllib.request.urlopen(url, timeout=60) as r:
        texto = r.read().decode("utf-8")
    j = json.loads(texto[texto.index("{"): texto.rindex("}") + 1])
    if j.get("status") == "error":
        raise RuntimeError(f"Erro ao ler a aba {nome}: {j.get('errors')}")
    cols = [key(c.get("label")) for c in j["table"]["cols"]]
    linhas = []
    for r in j["table"]["rows"]:
        cel = r.get("c") or []
        linhas.append({c: (cel[i].get("v") if i < len(cel) and cel[i] else "") for i, c in enumerate(cols)})
    return linhas


# ---------- quem é de qual segmento (mesma regra do filtro "Meu ano" do site) ----------
def ano_idx(v):
    k = key(v)
    n = re.search(r"(\d)", k)
    if not n:
        return -1
    return 8 + int(n[1]) if re.search(r"\bem\b|serie", k) else int(n[1]) - 1


def segmentos(o):
    a, b = ano_idx(o.get("do ano")), ano_idx(o.get("ate o ano"))
    if a >= 0 or b >= 0:
        a, b = (a if a >= 0 else 0), (b if b >= 0 else 11)
        faixas = {"EFAI": (0, 4), "EFAF": (5, 8), "EM": (9, 11)}
        return {s for s, (x, y) in faixas.items() if a <= y and b >= x}
    return {s for s in ("EFAI", "EFAF", "EM") if key(o.get(s.lower())) == "ok"}


# ---------- eventos ----------
def montar_eventos(olimpiadas, eventos, outras):
    oli = {key(o.get("sigla")): o for o in olimpiadas if norm(o.get("sigla"))}
    saida = []

    def add(sigla, rotulo, ini, fim, previsto, mes, detalhes, extra="", base_uid=""):
        o = oli.get(key(sigla), {})
        nome = norm(o.get("nome"))
        titulo = f"{sigla} · {rotulo}" if sigla else rotulo
        if mes:
            titulo += f" (previsto para {MESES[ini.month - 1]}, dia a confirmar)"
        elif previsto:
            titulo += " (previsto)"
        site = norm(o.get("site oficial") or o.get("site"))
        linhas = [nome,
                  "Data prevista com base no ano anterior: pode mudar. Confira sempre o site oficial." if previsto else "Data confirmada pelo organizador.",
                  extra, detalhes,
                  f"Site oficial: {site}" if re.match(r"^https?://", site, re.I) else "",
                  f"Todas as datas: {SITE}#{urllib.parse.quote(key(sigla))}" if sigla else f"Calendário completo: {SITE}"]
        saida.append({
            "uid": hashlib.sha1(f"{key(sigla)}|{base_uid or rotulo}|{ini.year}".encode()).hexdigest()[:16],
            "titulo": titulo, "ini": ini, "fim": fim or ini,
            "texto": "\n".join(x for x in linhas if x), "segs": segmentos(o) if sigla else {"EFAI", "EFAF", "EM"},
        })

    for r in eventos:
        sigla = norm(r.get("sigla"))
        if not sigla:
            continue
        sit = key(r.get("situacao")) or "previsto"
        if sit == "cancelado":
            continue
        conf = [x.strip() for x in key(r.get("etapas confirmadas")).split(",") if x.strip()]
        detalhes = norm(r.get("detalhes"))
        for nome, ci, cf, interno in ETAPAS:
            if interno:
                continue
            ini, fim = para_data(r.get(ci)), para_data(r.get(cf)) if cf else None
            if not ini and not fim:
                continue
            ok = sit == "confirmado" or key(nome) in conf
            previsto, mes = not ok, (not ok and "so o mes" in sit)
            if nome == "Inscrição":
                if ini and fim and fim != ini:
                    add(sigla, "Abrem as inscrições", ini, None, previsto, mes, detalhes,
                        f"Inscrições até {fim:%d/%m/%Y}.", "insc-ini")
                    add(sigla, "Último dia de inscrição", fim, None, previsto, mes, detalhes,
                        f"Inscrições desde {ini:%d/%m/%Y}.", "insc-fim")
                elif fim:
                    add(sigla, "Último dia de inscrição", fim, None, previsto, mes, detalhes, "", "insc-fim")
                else:
                    add(sigla, "Abrem as inscrições", ini, None, previsto, mes, detalhes, "", "insc-ini")
                continue
            ini = ini or fim
            if fim and fim > ini and (fim - ini).days > JANELA_CURTA:
                add(sigla, f"{nome} (início)", ini, None, previsto, mes, detalhes, f"Até {fim:%d/%m/%Y}.", nome + "-ini")
                add(sigla, f"{nome} (último dia)", fim, None, previsto, mes, detalhes, f"Desde {ini:%d/%m/%Y}.", nome + "-fim")
            else:
                add(sigla, nome, ini, fim if fim and fim > ini else None, previsto, mes, detalhes, "", nome)

    for r in outras:
        ini = para_data(r.get("data"))
        if not ini or key(r.get("visibilidade")) == "somente equipe":
            continue
        tipo = norm(r.get("tipo")) or "Outro"
        if key(tipo) in TIPOS_INTERNOS:
            continue
        fim = para_data(r.get("ate"))
        extra = " · ".join(x for x in (norm(r.get("horario")), norm(r.get("local / turma"))) if x)
        add(norm(r.get("sigla")), tipo, ini, fim if fim and fim > ini else None, False, False,
            norm(r.get("detalhes")), extra, f"{tipo}-{ini:%m%d}")
    return saida


# ---------- arquivo .ics ----------
def esc(s):
    return norm(s).replace("\\", "\\\\").replace(";", "\\;").replace(",", "\\,").replace("\r", "").replace("\n", "\\n")


def dobrar(linha):
    partes, atual = [], ""
    for ch in linha:
        if len((atual + ch).encode("utf-8")) > 74:
            partes.append(atual)
            atual = " "
        atual += ch
    partes.append(atual)
    return "\r\n".join(partes)


def gerar_ics(nome, eventos):
    linhas = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//CTO Mackenzie Brasilia//Calendario Olimpico//PT",
              "CALSCALE:GREGORIAN", "METHOD:PUBLISH", "X-WR-CALNAME:" + esc(nome),
              "X-WR-CALDESC:" + esc(f"Datas das olimpíadas do Centro de Treinamento Olímpico do Colégio Mackenzie Brasília. {SITE}"),
              "X-WR-TIMEZONE:America/Sao_Paulo", "REFRESH-INTERVAL;VALUE=DURATION:PT12H", "X-PUBLISHED-TTL:PT12H"]
    for e in sorted(eventos, key=lambda x: (x["ini"], x["titulo"])):
        linhas += ["BEGIN:VEVENT", "UID:" + e["uid"] + "@ctomackenziebrasilia.github.io",
                   "DTSTAMP:20260101T000000Z",
                   "DTSTART;VALUE=DATE:" + e["ini"].strftime("%Y%m%d"),
                   "DTEND;VALUE=DATE:" + (e["fim"] + timedelta(days=1)).strftime("%Y%m%d"),
                   "SUMMARY:" + esc(e["titulo"]), "DESCRIPTION:" + esc(e["texto"]), "URL:" + SITE,
                   "TRANSP:TRANSPARENT", "END:VEVENT"]
    linhas.append("END:VCALENDAR")
    return "\r\n".join(dobrar(l) for l in linhas) + "\r\n"


def main():
    if len(sys.argv) > 2 and sys.argv[1] == "--exemplo":
        d = json.loads(Path(sys.argv[2]).read_text(encoding="utf-8"))
        olimpiadas, eventos, outras = d["olimpiadas"], d["eventos"], d.get("outras", [])
    else:
        olimpiadas, eventos = ler_aba("Olimpiadas"), ler_aba("Eventos")
        try:
            outras = ler_aba("Treinos e outras datas")
        except Exception:
            outras = []
    if not olimpiadas or not eventos:
        raise SystemExit("Planilha vazia ou ilegível: nada foi alterado.")
    evs = montar_eventos(olimpiadas, eventos, outras)
    SAIDA.mkdir(exist_ok=True)
    for arq, (nome, seg) in AGENDAS.items():
        sel = [e for e in evs if seg in e["segs"]]
        (SAIDA / f"{arq}.ics").write_text(gerar_ics(nome, sel), encoding="utf-8", newline="")
        print(f"{arq}.ics: {len(sel)} eventos")
    # muda uma vez por mês: mantém o repositório ativo para o GitHub não pausar a automação
    (SAIDA / "mes.txt").write_text(date.today().strftime("%Y-%m") + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
