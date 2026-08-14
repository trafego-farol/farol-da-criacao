# -*- coding: utf-8 -*-
r"""
FAROL · EXTRAÇÃO VIA API DO ICLIPS
==================================
Substitui os dois exports manuais (lista_peças + relatório de atividades).

REGRAS DESTE SCRIPT — travadas por código:
  · SOMENTE LEITURA. Único método usado: GET.
  · Único endpoint consultado: /api/v1/projetos
  · NADA de financeiro: /financial/* está bloqueado por lista de proibição.
  · NUNCA cria, altera ou apaga nada no iClips.

COMO USAR (no seu PC, na pasta do Farol):
  1. Crie o arquivo  api_key.txt  com a chave dentro (uma linha, só a chave).
  2. py farol_api.py                 -> extrai os últimos 45 dias
     py farol_api.py 2026-01-01      -> extrai desde 01/01 (carga histórica)

  Gera em dados\:
     api_pauta.json      = pauta (peças em aberto, com prazo, DA e etapa)
     api_atividades.json = apontamentos (entregas, tempo, executor, refação)
  O farol_pasta.py passa a ler esses dois arquivos no lugar dos .xls/.xlsx.

SEGURANÇA: api_key.txt fica só na sua máquina. Não suba para o Netlify,
não mande por chat. Se a chave vazar, gere outra no iClips (Avatar > Chave de API).
"""
import os, sys, json, time, datetime, pathlib, collections
import urllib.request, urllib.error
try:
    import requests            # se instalado, é preferido: passa melhor pelo Cloudflare
    TEM_REQUESTS = True
except ImportError:
    TEM_REQUESTS = False

HERE  = pathlib.Path(__file__).parent
DADOS = HERE / "dados"; DADOS.mkdir(exist_ok=True)
BASE  = "https://public-api.iclips.com.br"

# ---- trava de segurança: nada além disto pode ser chamado ----
ENDPOINT_PERMITIDO = "/api/v1/projetos"
PROIBIDO = ("/financial", "/jobs", "/pieces", "/workflow-templates", "/piece-categories")

def carrega_chave():
    for nome in ("api_key.txt", "iclips_key.txt"):
        f = HERE / nome
        if f.exists():
            k = f.read_text(encoding="utf-8").strip()
            if k: return k
    k = os.environ.get("ICLIPS_KEY", "").strip()
    if k: return k
    print("Falta a chave. Crie o arquivo api_key.txt nesta pasta com a chave dentro.")
    sys.exit(1)

KEY = carrega_chave()

# O iClips fica atrás de Cloudflare, que recusa clientes sem identificação de navegador
# (erro 403 / "Error 1010"). Por isso mandamos cabeçalhos completos.
CABECALHOS = {
    "X-Api-Key": KEY,
    "Accept": "application/json, text/plain, */*",
    "Accept-Language": "pt-BR,pt;q=0.9,en;q=0.8",
    "Accept-Encoding": "identity",
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36"),
    "Referer": "https://app.iclips.com.br/",
    "Origin": "https://app.iclips.com.br",
    "Connection": "keep-alive",
}

def get(params, tent=1):
    """Único acesso à API. GET, endpoint fixo, com respeito ao limite de 10 req/min."""
    path = ENDPOINT_PERMITIDO
    assert not any(p in path for p in PROIBIDO), "endpoint bloqueado"
    url = BASE + path + "?" + "&".join(f"{k}={v}" for k, v in params.items())
    if TEM_REQUESTS:
        try:
            r = requests.get(url, headers=CABECALHOS, timeout=120)
            if r.status_code == 200: return r.json()
            if r.status_code == 429 and tent <= 5:
                esp = int(r.headers.get("Retry-After", 30))
                print(f"      [limite de taxa] aguardando {esp}s"); time.sleep(esp + 2)
                return get(params, tent + 1)
            if r.status_code == 401:
                print("      ERRO 401: chave inválida, expirada ou plano não é PRO."); sys.exit(1)
            if r.status_code == 403:
                print("      ERRO 403: bloqueado pelo Cloudflare antes de chegar na API.")
                print("      O IP da sua rede pode precisar ser liberado junto ao suporte do iClips.")
                sys.exit(1)
            print(f"      ERRO HTTP {r.status_code}: {r.text[:200]}"); return None
        except requests.RequestException as e:
            if tent <= 3: time.sleep(5); return get(params, tent + 1)
            print(f"      ERRO: {e}"); return None
    req = urllib.request.Request(url, headers=CABECALHOS)
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        if e.code == 429 and tent <= 5:
            esp = int(e.headers.get("Retry-After", 30))
            print(f"      [limite de taxa] aguardando {esp}s"); time.sleep(esp + 2)
            return get(params, tent + 1)
        if e.code == 401:
            print("      ERRO 401: chave inválida, expirada ou plano não é PRO."); sys.exit(1)
        corpo = e.read().decode("utf-8", "ignore")
        if e.code == 403 and ("1010" in corpo or "cloudflare" in corpo.lower()):
            print("      ERRO 403 (Cloudflare): a requisição foi bloqueada antes de chegar na API.")
            print("      Tente instalar a biblioteca requests, que costuma passar:")
            print("          pip install requests")
            print("      e rode de novo. Se persistir, o IP da agência pode precisar ser liberado no iClips.")
            sys.exit(1)
        print(f"      ERRO HTTP {e.code}: {corpo[:200]}")
        return None
    except Exception as e:
        if tent <= 3:
            time.sleep(5); return get(params, tent + 1)
        print(f"      ERRO: {e}"); return None

def janelas(ini, fim, dias=30):
    """A API aceita no máximo 31 dias por consulta."""
    a = ini
    while a <= fim:
        b = min(a + datetime.timedelta(days=dias - 1), fim)
        yield a, b
        a = b + datetime.timedelta(days=1)

def extrai(ini, fim):
    projetos, req = [], 0
    for a, b in janelas(ini, fim):
        print(f"   janela {a} → {b}")
        pag = 1
        while True:
            r = get({"dataInicio": f"{a}T00:00:00", "dataFim": f"{b}T23:59:59",
                     "page": pag, "pageSize": 50})
            req += 1
            if not r: break
            d = r.get("data", []) or []
            projetos += d
            meta = r.get("meta", {}) or {}
            tot = meta.get("totalCount", 0)
            print(f"      página {pag}: {len(d)} projetos (total no período: {tot})")
            if len(d) < 50 or pag * 50 >= tot: break
            pag += 1
            time.sleep(6.5)          # 10 req/min com folga
        time.sleep(6.5)
    print(f"   {len(projetos)} projetos · {req} requisições")
    return projetos

def hhmm_para_min(s):
    if not s: return 0
    try:
        p = str(s).split(":")
        return int(p[0]) * 60 + int(p[1]) + (int(p[2]) / 60 if len(p) > 2 else 0)
    except Exception: return 0

def transforma(projetos):
    """Converte a hierarquia da API nas duas estruturas que o Farol usa."""
    pauta, ativ = [], []
    ABERTO = {"concl", "final", "entreg", "arquiv", "cancel"}   # status que NÃO é pauta aberta
    for p in projetos:
        cli  = (p.get("cliente") or {}).get("nome", "") or ""
        grp  = p.get("grupoCliente")
        grp  = (grp or {}).get("nome", "") if isinstance(grp, dict) else (grp or "")
        proj = p.get("nomeProjeto", "") or ""
        dts  = p.get("datas") or {}
        for pc in p.get("pecas", []) or []:
            nome  = pc.get("nomePeca", "") or ""
            tit   = pc.get("tituloAtividade", "") or nome
            prazo = (pc.get("fimPlanejado") or "")[:10]
            wfs   = pc.get("workflows", []) or []
            # etapa atual = último workflow sem fim; se todos fecharam, o último
            atual, resp = "", []
            abertos = [w for w in wfs if not w.get("fim")]
            wref = (abertos[-1] if abertos else (wfs[-1] if wfs else None))
            if wref:
                atual = wref.get("nome", "") or ""
                for a in wref.get("atividades", []) or []:
                    ex = (a.get("executor") or {}).get("nome")
                    if ex and ex not in resp: resp.append(ex)
            if not resp:      # sem executor na etapa atual: pega de qualquer etapa
                for w in wfs:
                    for a in w.get("atividades", []) or []:
                        ex = (a.get("executor") or {}).get("nome")
                        if ex and ex not in resp: resp.append(ex)
            if abertos or not wfs:
                pauta.append({"cliente": cli, "grupo": grp, "projeto": proj,
                              "idJobPeca": pc.get("idJobPeca"), "peca": tit, "fmt": nome,
                              "prazo": prazo, "status": atual, "das": resp,
                              "conclusaoPrevista": (dts.get("conclusaoEstimada") or "")[:10]})
            for w in wfs:
                etapa = w.get("nome", "") or ""
                for a in w.get("atividades", []) or []:
                    ex  = a.get("executor") or {}
                    fim = a.get("fimPlay") or a.get("inicioPlay") or ""
                    if not fim: continue
                    ativ.append({"da": ex.get("nome", "") or "", "dia": fim[:10], "fim": fim[11:16],
                                 "cliente": cli, "grupo": grp, "projeto": proj,
                                 "idJobPeca": pc.get("idJobPeca"), "peca": tit, "fmt": nome,
                                 "etapa": etapa, "refacao": bool(w.get("refacao")),
                                 "dep": ex.get("departamento", "") or "",
                                 "min": round(hhmm_para_min(a.get("tempoGasto")), 1)})
    return pauta, ativ

def main():
    fim = datetime.date.today()
    ini = datetime.date.fromisoformat(sys.argv[1]) if len(sys.argv) > 1 else fim - datetime.timedelta(days=45)
    print("=" * 72)
    print(f"FAROL · extração via API  ·  entrada de job entre {ini} e {fim}")
    print("somente leitura · endpoint /api/v1/projetos · sem dados financeiros")
    print("=" * 72)
    t0 = time.time()
    projetos = extrai(ini, fim)
    if not projetos:
        print("Nada retornado. Confira a chave e o período."); sys.exit(1)
    pauta, ativ = transforma(projetos)
    json.dump(pauta, open(DADOS / "api_pauta.json", "w", encoding="utf-8"), ensure_ascii=False)
    json.dump(ativ,  open(DADOS / "api_atividades.json", "w", encoding="utf-8"), ensure_ascii=False)
    hoje = fim.isoformat()
    print("\n" + "=" * 72)
    print(f"PAUTA ....... {len(pauta)} peças em aberto · {sum(1 for x in pauta if x['prazo']==hoje)} vencem hoje")
    print(f"ATIVIDADES .. {len(ativ)} apontamentos · {len({a['dia'] for a in ativ})} dias")
    print(f"              {sum(1 for a in ativ if a['dia']==hoje)} apontamentos hoje")
    print(f"              {sum(1 for a in ativ if a['refacao'])} marcados como refação pela própria API")
    deps = collections.Counter(a["dep"] for a in ativ if a["dep"])
    if deps: print(f"DEPARTAMENTOS {dict(deps.most_common(8))}")
    print(f"\ngravado em dados\\api_pauta.json e dados\\api_atividades.json")
    print(f"tempo: {time.time()-t0:.0f}s")
    print("=" * 72)

if __name__ == "__main__":
    main()
