# -*- coding: utf-8 -*-
"""
Análise de pauta por CLASSE DE CLIENTE (AA/A/B/C).
Não mexe no dashboard: lê a pauta do iClips e devolve a ordem de atendimento.

uso:  py analise_pauta_classe.py [caminho\\lista_pecas.xlsx]
      sem argumento, usa dados\\lista_pecas.xlsx
"""
import json, re, sys, unicodedata, pathlib, collections
import openpyxl

HERE = pathlib.Path(__file__).parent
BASE = json.load(open(HERE/'clientes_classes.json', encoding='utf-8'))
CLASSE, APELIDOS, GRUPOS, PESO = BASE['CLASSE'], BASE['APELIDOS'], BASE['GRUPOS'], BASE['PESO']

def norm(s):
    t = unicodedata.normalize('NFKD', str(s or '')).encode('ascii','ignore').decode().upper()
    t = re.sub(r'\[[^\]]*\]', '', t)
    t = re.sub(r'\b(LTDA|SPE|S/A|SA|ME|EIRELI)\b', '', t)
    t = re.sub(r'[^A-Z0-9 ]', ' ', t)
    return re.sub(r'\s+', ' ', t).strip()

NC = {norm(k): v for k, v in CLASSE.items()}
NA = {norm(k): norm(v) for k, v in APELIDOS.items()}
GRP = {}
for g, ms in GRUPOS.items():
    for m in ms: GRP[norm(m)] = g

def classe_de(cliente):
    """Devolve (classe, grupo economico) para o nome do cliente como vem do iClips."""
    n = norm(cliente)
    n = NA.get(n, n)                                   # resolve apelido
    if n in NC: return NC[n], GRP.get(n, '')
    cand = [k for k in NC if k and (k in n or n in k)]  # casamento parcial
    if cand:
        cand.sort(key=len, reverse=True)
        return NC[cand[0]], GRP.get(cand[0], '')
    return '?', ''                                     # não cadastrado

def ler_pauta(arq):
    ws = openpyxl.load_workbook(arq, data_only=True).active
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        proj = str(r[0] or '').replace('\xa0', ' ')
        ati  = str(r[1] or '').replace('\xa0', ' ')
        cli  = proj.split('  ·  ')[-1] if '  ·  ' in proj else proj
        peca = ati.split('  ·  ')[-1]
        fmt  = re.sub(r'^\[[^\]]*\]\s*', '', ati.split('  ·  ')[0]).strip()
        das  = [x.strip() for x in str(r[8] or '').split(', ') if x.strip()]
        out.append({'cliente': cli.strip(), 'peca': peca.strip(), 'fmt': fmt,
                    'prazo': str(r[5])[:10], 'status': str(r[9] or ''), 'das': das})
    return out

def main(arq=None):
    arq = pathlib.Path(arq) if arq else HERE/'dados'/'lista_pecas.xlsx'
    P = ler_pauta(arq)
    for p in P:
        p['classe'], p['grupo'] = classe_de(p['cliente'])
        p['peso'] = PESO.get(p['classe'], -1)

    print(f"\nPAUTA: {len(P)} peças · fonte {arq.name}\n" + "="*84)

    print("\n1) PEÇAS POR CLASSE DE CLIENTE")
    c = collections.Counter(p['classe'] for p in P)
    for k in ['AA','A','B','C','—','?']:
        if c.get(k): print(f"   {k:3s} {c[k]:4d} peças  ({round(100*c[k]/len(P))}%)")

    nc = sorted({p['cliente'] for p in P if p['classe'] == '?'})
    if nc:
        print(f"\n   ATENÇÃO: {len(nc)} cliente(s) sem classe no cadastro:")
        for x in nc: print(f"      · {x}")

    print("\n2) ORDEM DE ATENDIMENTO — clientes da pauta, por prioridade")
    porcli = collections.defaultdict(lambda: {'n':0,'cl':'','gr':''})
    for p in P:
        d = porcli[p['cliente']]; d['n'] += 1; d['cl'] = p['classe']; d['gr'] = p['grupo']
    ordem = sorted(porcli.items(), key=lambda kv: (-PESO.get(kv[1]['cl'], -1), -kv[1]['n']))
    for i, (cli, d) in enumerate(ordem, 1):
        g = f"  [{d['gr']}]" if d['gr'] else ''
        print(f"   {i:3d}. {d['cl']:3s} {d['n']:3d} peças · {cli}{g}")

    print("\n3) GRUPOS ECONÔMICOS NA PAUTA — juntar na mesma janela")
    pg = collections.defaultdict(lambda: {'n':0,'cl':set(),'cli':set()})
    for p in P:
        if p['grupo']:
            g = pg[p['grupo']]; g['n'] += 1; g['cl'].add(p['classe']); g['cli'].add(p['cliente'])
    for g, d in sorted(pg.items(), key=lambda kv: -kv[1]['n']):
        print(f"   {g:20s} {d['n']:3d} peças · {len(d['cli'])} empresa(s) · classe {'/'.join(sorted(d['cl']))}")

    print("\n4) CARGA POR PROFISSIONAL, COM PESO DE CLIENTE")
    pp = collections.defaultdict(lambda: collections.Counter())
    for p in P:
        for da in p['das']: pp[da][p['classe']] += 1
    linhas = []
    for da, cc in pp.items():
        tot = sum(cc.values()); pes = sum(PESO.get(k, 0)*v for k, v in cc.items())
        linhas.append((pes, tot, da, cc))
    for pes, tot, da, cc in sorted(linhas, reverse=True):
        det = ' · '.join(f"{k}:{cc[k]}" for k in ['AA','A','B','C','—','?'] if cc.get(k))
        print(f"   {da[:34]:36s} {tot:3d} peças · peso {pes:3d}  ({det})")

    print("\n" + "="*84)
    aa = [p for p in P if p['classe'] == 'AA']
    print(f"RESUMO: {len(aa)} peças de clientes AA exigem atenção primeiro "
          f"({round(100*len(aa)/len(P))}% da pauta), distribuídas em "
          f"{len({p['cliente'] for p in aa})} clientes.")

if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else None)
