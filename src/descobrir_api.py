# -*- coding: utf-8 -*-
r"""
DESCOBRIR A CONEXÃO — testa combinações até achar a que funciona.
Somente leitura. Rode:  py descobrir_api.py

Sua agência usa a Content Data API:
   https://content-data.iclips.com.br/clients/ICLIPSID/content-data?...
O ICLIPSID é um código que o iClips envia por e-mail. Se você não tiver,
este script tenta adivinhar a partir da sua chave.
"""
import os, sys, json, pathlib, datetime
try:
    import requests; USA_REQ = True
except ImportError:
    import urllib.request, urllib.error; USA_REQ = False

HERE = pathlib.Path(__file__).parent

def chave():
    for n in ("api_key.txt", "iclips_key.txt"):
        f = HERE/n
        if f.exists():
            k = f.read_text(encoding="utf-8").strip()
            if k and "COLE A CHAVE" not in k.upper(): return k
    return os.environ.get("ICLIPS_KEY", "").strip()

KEY = chave()
if not KEY:
    print("Sem chave. Preencha o api_key.txt."); sys.exit(1)

ID_ARQ = HERE/"iclips_id.txt"
ID_MANUAL = ID_ARQ.read_text(encoding="utf-8").strip() if ID_ARQ.exists() else ""

NAV = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
       "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36")
hoje = datetime.date.today(); ini = hoje - datetime.timedelta(days=3)

def bate(url, headers):
    try:
        if USA_REQ:
            r = requests.get(url, headers=headers, timeout=45)
            return r.status_code, r.text[:400]
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=45) as r:
            return r.status, r.read().decode("utf-8", "ignore")[:400]
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", "ignore")[:300]
    except Exception as e:
        return 0, f"{type(e).__name__}: {str(e)[:120]}"

# candidatos de ICLIPSID: o do arquivo, o número dentro da chave, e a chave inteira
cands = []
if ID_MANUAL: cands.append(("do arquivo iclips_id.txt", ID_MANUAL))
partes = KEY.split("_")
if len(partes) >= 3 and partes[2].isdigit(): cands.append(("número dentro da chave", partes[2]))
cands.append(("chave inteira", KEY))

# nomes de cabeçalho possíveis
heads = [("platform", {"platform": KEY}),
         ("X-Api-Key", {"X-Api-Key": KEY}),
         ("Authorization Bearer", {"Authorization": f"Bearer {KEY}"}),
         ("token", {"token": KEY})]

base = "https://content-data.iclips.com.br/clients/{}/content-data"
qs   = f"?rangeStart={ini}&rangeEnd={hoje}&limit=5&page=1"

print("="*70); print("PROCURANDO A COMBINAÇÃO QUE FUNCIONA"); print("="*70)
print(f"chave: {KEY[:18]}...{KEY[-6:]}\n")
achou = None
for rot_id, cid in cands:
    for rot_h, h in heads:
        hh = dict(h); hh.update({"Accept": "application/json", "User-Agent": NAV})
        cod, corpo = bate(base.format(cid) + qs, hh)
        marca = "OK  " if cod == 200 else "    "
        print(f"  {marca}[{cod or 'XX'}] id={rot_id:26s} header={rot_h}")
        if cod == 200:
            achou = (cid, rot_h, hh, corpo); break
    if achou: break

print("\n" + "="*70)
if achou:
    cid, rot_h, hh, corpo = achou
    print("FUNCIONOU!")
    print(f"  ICLIPSID .... {cid}")
    print(f"  cabeçalho ... {rot_h}")
    try:
        d = json.loads(corpo)
        regs = d.get("data", d if isinstance(d, list) else [])
        print(f"  registros ... {len(regs)} na amostra")
        if regs:
            print("\n  campos disponíveis:")
            for k in sorted(regs[0].keys())[:40]: print(f"     {k}")
    except Exception:
        print(f"  amostra: {corpo[:200]}")
    (HERE/"api_conexao.json").write_text(json.dumps(
        {"iclipsid": cid, "header": rot_h}, ensure_ascii=False, indent=1), encoding="utf-8")
    print("\n  gravei api_conexao.json — me mande esse arquivo.")
else:
    print("NENHUMA COMBINAÇÃO FUNCIONOU.")
    print("\nProvável causa: falta o ICLIPSID, que é um código no formato")
    print("   709762a1-f1cz-1d31-3134-31zap3c531tg")
    print("e que o iClips envia POR E-MAIL quando libera a API.")
    print("\nO que fazer:")
    print("  1. Procure no seu e-mail por 'ICLIPSID' ou 'Content Data API'.")
    print("  2. Se achar, crie o arquivo  iclips_id.txt  nesta pasta com o código")
    print("     dentro e rode este script de novo.")
    print("  3. Se não achar, peça ao suporte do iClips o seu ICLIPSID.")
print("="*70)
