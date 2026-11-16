# Gera uma página de entrada por olimpíada em o/<slug>/ (prévia própria no WhatsApp)
# a partir da aba Olimpiadas da planilha pública. Roda no GitHub Actions (diariamente).
# Cuidados contra prévia errada:
#  - só usa Sigla e Nome (sem datas, que mudam);
#  - a pasta é refeita do zero a cada execução, então sigla renomeada ou removida some;
#  - siglas que dariam o mesmo endereço são puladas (nenhuma das duas ganha página);
#  - se a planilha não puder ser lida ou vier quase vazia, nada é apagado.
import csv, html, io, json, os, shutil, sys, textwrap, unicodedata, urllib.request, re
from PIL import Image, ImageDraw, ImageFont

PLANILHA = "1wF10E2_omNchLjgwAkMNBRixaEMf1wp17DaxzZsgKag"
SITE = "https://ctomackenziebrasilia.github.io/"
RAIZ = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
AQUI = os.path.dirname(os.path.abspath(__file__))
DESTINO = os.path.join(RAIZ, "o")
MINIMO = 10


def chave(s):  # igual ao key() do site
    s = (s or "").strip().lower()
    return "".join(c for c in unicodedata.normalize("NFD", s) if not unicodedata.combining(c))


def slug(s):  # igual ao slugDe() do site
    return re.sub(r"[^a-z0-9]+", "-", chave(s)).strip("-")


def ler_planilha(arquivo=None):
    if arquivo:
        dados = open(arquivo, encoding="utf-8").read()
    else:
        url = f"https://docs.google.com/spreadsheets/d/{PLANILHA}/gviz/tq?tqx=out:csv&sheet=Olimpiadas"
        dados = urllib.request.urlopen(url, timeout=60).read().decode("utf-8")
    linhas = list(csv.DictReader(io.StringIO(dados)))
    out = []
    for l in linhas:
        sig = (l.get("Sigla") or "").strip()
        nome = (l.get("Nome") or "").strip()
        if sig:
            out.append((sig, nome or sig))
    return out


def fonte(nome, tam):
    return ImageFont.truetype(os.path.join(AQUI, nome), tam)


def imagem(sig, nome, caminho):
    im = Image.open(os.path.join(AQUI, "base.png")).convert("RGB")
    d = ImageDraw.Draw(im)
    larg = 1200 - 72 * 2
    tam = 112
    while tam > 56 and d.textlength(sig, font=fonte("Inter-ExtraBold.otf", tam)) > larg:
        tam -= 4
    d.text((72, 240), sig, font=fonte("Inter-ExtraBold.otf", tam), fill="white")
    y = 240 + tam + 18
    for t in (38, 34, 30):
        f = fonte("Inter-Medium.otf", t)
        cols = max(10, int(larg / (t * 0.52)))
        linhas = textwrap.wrap(nome, cols)
        if len(linhas) <= 2 and all(d.textlength(x, font=f) <= larg for x in linhas):
            break
    if len(linhas) > 2:  # nome muito longo: corta com reticências em vez de sumir sem aviso
        linhas = linhas[:2]
        while d.textlength(linhas[1] + "…", font=f) > larg:
            linhas[1] = linhas[1][:-1]
        linhas[1] = linhas[1].rstrip() + "…"
    for x in linhas:
        d.text((72, y), x, font=f, fill=(255, 255, 255, 235))
        y += int(t * 1.25)
    im.save(caminho, "JPEG", quality=86, optimize=True)


def pagina(sig, nome, s):
    k = chave(sig)
    alvo = SITE + "#" + urllib.parse.quote(k, safe="")
    url = f"{SITE}o/{s}/"
    titulo = html.escape(f"{sig} – {nome}" if nome != sig else sig)
    desc = "Datas, site oficial e como participar, no Calendário Olímpico do CTO Mackenzie Brasília."
    return f"""<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<title>{titulo} – Calendário Olímpico CTO</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex">
<meta name="cto-sigla" content="{html.escape(k)}">
<meta property="og:type" content="website">
<meta property="og:locale" content="pt_BR">
<meta property="og:site_name" content="CTO – Colégio Mackenzie Brasília">
<meta property="og:title" content="{titulo}">
<meta property="og:description" content="{desc}">
<meta property="og:url" content="{url}">
<meta property="og:image" content="{url}og.jpg">
<meta property="og:image:secure_url" content="{url}og.jpg">
<meta property="og:image:type" content="image/jpeg">
<meta property="og:image:width" content="1200">
<meta property="og:image:height" content="630">
<meta property="og:image:alt" content="{titulo} no Calendário Olímpico do CTO">
<meta name="twitter:card" content="summary_large_image">
<meta name="theme-color" content="#A8192E">
<link rel="canonical" href="{html.escape(alvo)}">
<meta http-equiv="refresh" content="0; url={html.escape(alvo)}">
<script>location.replace({json.dumps(alvo)});</script>
<style>body{{font-family:system-ui,sans-serif;background:#A8192E;color:#fff;display:grid;place-items:center;min-height:100vh;margin:0;text-align:center}}a{{color:#fff}}</style>
</head>
<body><p>Abrindo <a href="{html.escape(alvo)}">{titulo}</a> no Calendário Olímpico…</p></body>
</html>
"""


def main():
    import urllib.parse  # noqa
    arquivo = sys.argv[1] if len(sys.argv) > 1 else None
    try:
        olis = ler_planilha(arquivo)
    except Exception as e:
        print("Não foi possível ler a planilha; nada foi alterado.", e)
        return 0
    if len(olis) < MINIMO:
        print(f"Planilha com só {len(olis)} olimpíadas; nada foi alterado.")
        return 0
    por_slug = {}
    for sig, nome in olis:
        por_slug.setdefault(slug(sig), []).append((sig, nome))
    tmp = DESTINO + ".novo"
    shutil.rmtree(tmp, ignore_errors=True)
    os.makedirs(tmp)
    feitas, puladas = [], []
    for s, itens in sorted(por_slug.items()):
        if not s or len(itens) > 1:
            puladas += [i[0] for i in itens]
            continue
        sig, nome = itens[0]
        os.makedirs(os.path.join(tmp, s))
        open(os.path.join(tmp, s, "index.html"), "w", encoding="utf-8").write(pagina(sig, nome, s))
        imagem(sig, nome, os.path.join(tmp, s, "og.jpg"))
        feitas.append(s)
    open(os.path.join(tmp, "lista.json"), "w", encoding="utf-8").write(json.dumps(feitas))
    shutil.rmtree(DESTINO, ignore_errors=True)
    os.rename(tmp, DESTINO)
    print(f"{len(feitas)} prévias geradas.")
    if puladas:
        print("Sem prévia própria (siglas que dariam o mesmo endereço):", ", ".join(puladas))
    return 0


if __name__ == "__main__":
    import urllib.parse
    sys.exit(main())
