# -*- coding: utf-8 -*-
r"""
DIAGNÓSTICO — descobre por que a API está bloqueando.
Testa 4 formas diferentes de chamar e diz qual funciona.
Somente leitura. Rode:  py diagnostico_api.py
"""
import os, sys, json, pathlib, datetime
HERE = pathlib.Path(__file__).parent
def chave():
    for n in ("api_key.txt","iclips_key.txt"):
        f=HERE/n
        if f.exists():
            k=f.read_text(encoding="utf-8").strip()
            if k and "COLE A CHAVE" not in k: return k
    return os.environ.get("ICLIPS_KEY","").strip()
KEY=chave()
if not KEY: print("Sem chave. Preencha api_key.txt."); sys.exit(1)
print(f"chave encontrada: {KEY[:18]}...{KEY[-6:]}  ({len(KEY)} caracteres)\n")

hoje=datetime.date.today(); ini=hoje-datetime.timedelta(days=3)
URL=("https://public-api.iclips.com.br/api/v1/projetos"
     f"?dataInicio={ini}T00:00:00&dataFim={hoje}T23:59:59&page=1&pageSize=5")
NAV=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
     "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36")

def mostra(nome, cod, corpo):
    if cod==200:
        try: n=len(json.loads(corpo).get("data",[]))
        except Exception: n="?"
        print(f"  [OK]   {nome}  ->  200, {n} projetos")
        return True
    m=""
    if "1010" in str(corpo): m=" (Cloudflare 1010 - assinatura do cliente barrada)"
    elif cod==401: m=" (chave inválida ou plano não é PRO)"
    print(f"  [FALHA] {nome}  ->  {cod}{m}")
    return False

ok=[]
# 1 urllib sem cabeçalho de navegador
import urllib.request, urllib.error
def tenta_urllib(nome, headers):
    try:
        r=urllib.request.urlopen(urllib.request.Request(URL,headers=headers),timeout=60)
        return mostra(nome, r.status, r.read().decode("utf-8","ignore"))
    except urllib.error.HTTPError as e:
        return mostra(nome, e.code, e.read().decode("utf-8","ignore")[:300])
    except Exception as e:
        print(f"  [ERRO]  {nome}  ->  {type(e).__name__}: {str(e)[:90]}"); return False

print("testando 4 formas de conexão:\n")
ok.append(("urllib simples", tenta_urllib("urllib simples", {"X-Api-Key":KEY,"Accept":"application/json"})))
ok.append(("urllib + navegador", tenta_urllib("urllib + cabeçalho de navegador",
    {"X-Api-Key":KEY,"Accept":"application/json","User-Agent":NAV,
     "Accept-Language":"pt-BR,pt;q=0.9","Referer":"https://app.iclips.com.br/"})))
try:
    import requests
    try:
        r=requests.get(URL,headers={"X-Api-Key":KEY,"Accept":"application/json","User-Agent":NAV},timeout=60)
        ok.append(("requests", mostra("requests + navegador", r.status_code, r.text[:300])))
    except Exception as e:
        print(f"  [ERRO]  requests  ->  {str(e)[:90]}")
except ImportError:
    print("  [--]    requests  ->  não instalado (pip install requests)")
import subprocess, shutil
if shutil.which("curl"):
    try:
        p=subprocess.run(["curl","-s","-o","-","-w","\n%{http_code}","-H",f"X-Api-Key: {KEY}",
                          "-H","Accept: application/json","-H",f"User-Agent: {NAV}",URL],
                         capture_output=True,text=True,timeout=60)
        linhas=p.stdout.rsplit("\n",1); cod=int(linhas[-1]) if linhas[-1].strip().isdigit() else 0
        ok.append(("curl", mostra("curl do Windows", cod, linhas[0][:300])))
    except Exception as e:
        print(f"  [ERRO]  curl  ->  {str(e)[:90]}")
else:
    print("  [--]    curl  ->  não encontrado")

print("\n" + "="*64)
if any(v for _,v in ok): print("Pelo menos uma forma funcionou. Me diga qual e eu ajusto o robô.")
else:
    print("Todas bloquearam. Provável causa: o Cloudflare do iClips está barrando")
    print("o IP ou a rede da agência para chamadas fora do navegador.")
    print("Peça ao suporte do iClips para liberar o acesso via API para o seu IP.")
print("="*64)
