# -*- coding: utf-8 -*-
"""
FAROL PASTA — monta o dashboard a partir dos 2 relatórios do iClips e publica no Netlify.
Fonte (coloque os arquivos na subpasta 'dados'):
  • Gestão de Atividades (período = hoje, colaborador = Criação/Redatores/Vídeo)  -> lista_peças*.xlsx
  • Atividades Realizadas por Colaborador (período amplo, ex.: 01/01 até hoje)     -> relatorio_atividade*.xls
Usa SEMPRE o arquivo mais recente de cada tipo. Lógica idêntica à validada no dashboard.
Requisitos: pip install openpyxl pandas lxml requests
Rode a cada 30 min pelo Agendador de Tarefas (aponte para Farol_Pasta.bat).
"""
import json, re, sys, datetime, unicodedata, hashlib, pathlib, glob, statistics
from collections import defaultdict
try:
    import openpyxl, pandas as pd, requests
except ImportError:
    sys.exit("Instale: pip install openpyxl pandas lxml requests")

HERE=pathlib.Path(__file__).parent
DADOS=HERE/"dados"; DADOS.mkdir(exist_ok=True)
cfg=json.loads((HERE/"farol_config.json").read_text(encoding="utf-8"))
K=json.loads((HERE/"farol_constantes.json").read_text(encoding="utf-8"))
TPL=(HERE/"farol_template_deploy.html").read_text(encoding="utf-8")
SNAP=HERE/"snapshot_pauta.json"
hoje=datetime.date.today().isoformat()

ALOC={k:tuple(v) for k,v in K["ALOC"].items()}
TEAM_RED=set(K["TEAM_RED"]); TEAM_VID=set(K["TEAM_VID"]); EXCLUIR=set(K["EXCLUIR"])
SAIDA=K.get("SAIDA",{})
AUSENTES=K.get("AUSENTES",{})   # dia -> pessoas fora (folga, atestado, assunto pessoal): não entram na capacidade do dia   # pessoa -> data a partir da qual não conta mais (histórico anterior é preservado)
HOUSE={n for n,v in ALOC.items() if v[0]=="House"}
tv=set(TEAM_RED)|set(TEAM_VID)|{'TRÁFEGO'}
SEP='\xa0\xa0·\xa0\xa0'
TIPO_MAP={'SOCIAL MEDIA':'Social Media','OFF':'Off','INBOUND':'Inbound','DEV':'Dev / On','MKT DIGITAL':'Mídia Paga','SCIENCE SEO':'SEO'}
SM={'arte':'Arte','criação de arte':'Arte','criação':'Arte','arte/figma':'Arte','arte final':'Arte final','novo':'Novo (fila)','refação interna':'Refação interna','refeção cliente':'Refação cliente','refação cliente':'Refação cliente','ajuste interno':'Ajuste interno','ajuste':'Ajuste interno','ajuste cliente':'Ajuste cliente','ajuste cliente arte':'Ajuste cliente','cartelas':'Cartelas','planejamento':'Planejamento','diagramar e-mail':'Diagramação e-mail','salvar peças no drive':'Salvar peças no Drive','aprovação diretor de criação':'Aprovação Dir. Criação','aprovação interna':'Aprovação interna','edição video':'Edição de vídeo'}
ns=lambda s: SM.get((s or '').strip().casefold(),(s or '').strip() or '(sem status)')
_n=lambda s:unicodedata.normalize('NFD',(s or '').upper()).encode('ascii','ignore').decode()
is_house=lambda c: 'NOVA ERA' in _n(c) or 'PATIO GOURMET' in _n(c)

def newest(pat):
    fs=sorted(DADOS.glob(pat), key=lambda p:p.stat().st_mtime)
    return fs[-1] if fs else None

# FONTE DOS DADOS — prioriza a API (dados/api_*.json); usa os exports manuais como reserva
API_PAUTA = DADOS/"api_pauta.json"
API_ATIV  = DADOS/"api_atividades.json"
VIA_API   = API_PAUTA.exists() and API_ATIV.exists()
f_pauta = API_PAUTA if VIA_API else (newest("lista*.xlsx") or newest("*.xlsx"))
f_ativ  = sorted(DADOS.glob("relatorio*.xls"), key=lambda p:p.stat().st_mtime)   # em modo API viram histórico complementar
if not f_pauta: sys.exit("Coloque o export da Gestão de Atividades (lista_peças*.xlsx) na pasta 'dados'.")

# ---------- limpeza de título (ETL) ----------
def limpa_titulo(s):
    """Remove sujeira do nome da peça: tag [TIPO], {datas}, (ESCOPO), separadores soltos."""
    t=str(s or '')
    t=re.sub(r'^\s*\[[^\]]*\]\s*','',t)          # tag do tipo no começo
    t=re.sub(r'\{[^}]*\}','',t)                   # {29/07/2026}
    t=re.sub(r'\((?:ESCOPO|EXTRA)\)','',t,flags=re.I)
    t=t.replace('\xa0',' ')
    t=re.sub(r'^[\s\-–—|·:]+','',t)               # lixo no início
    t=re.sub(r'[\s\-–—|·:]+$','',t)               # lixo no fim
    t=re.sub(r'\s{2,}',' ',t)
    t=re.sub(r'\s*\|\s*\|\s*',' | ',t)
    return t.strip() or str(s or '').strip()

# ---------- PAUTA ----------
def _prazo(v):
    """O export do iClips às vezes traz a data como valor, às vezes como texto."""
    if isinstance(v,datetime.datetime): return v.date().isoformat()
    if isinstance(v,datetime.date):     return v.isoformat()
    t=str(v or '').strip()
    if not t or t.lower()=='none': return None
    m=re.match(r'(\d{2})/(\d{2})/(\d{4})',t)          # dd/mm/aaaa
    if m: return f'{m.group(3)}-{m.group(2)}-{m.group(1)}'
    m=re.match(r'(\d{4})-(\d{2})-(\d{2})',t)          # aaaa-mm-dd
    if m: return t[:10]
    return None

def load_pauta_api(path, inc_house=False):
    """Lê dados/api_pauta.json (extraído pela API do iClips). Mesmo formato de saída do .xlsx."""
    out={}
    for r in json.load(open(path,encoding='utf-8')):
        cli=(r.get('cliente') or '').strip()
        if is_house(cli) and not inc_house: continue
        das=[d.strip() for d in (r.get('das') or []) if d.strip() not in EXCLUIR and d.strip() not in SAIDA]
        if not das: continue
        nome=str(r.get('fmt') or '')
        m=re.match(r'\s*\[([^\]]+)\]\s*(.*)',nome)
        tipo=TIPO_MAP.get(m.group(1).strip().upper(),'Outros') if m else 'Outros'
        fmt=((m.group(2) if m else nome) or '').strip().upper() or '(SEM FORMATO)'
        k=(str(r.get('projeto') or '')+'|'+str(r.get('idJobPeca') or r.get('peca')))
        out[k]={'cliente':cli,'peca':limpa_titulo(r.get('peca')),'tipo':tipo,'fmt':fmt,
                'prio':'','das':das,'prazo':(r.get('prazo') or None),'status':ns(r.get('status') or '')}
    return out

def load_pauta(path, inc_house=False):
    if str(path).lower().endswith('.json'): return load_pauta_api(path, inc_house)
    wb=openpyxl.load_workbook(path,data_only=True); out={}
    for r in wb.active.iter_rows(min_row=2, values_only=True):
        if not r[0]: continue
        pp=str(r[0]).split(SEP); cli=pp[-1].strip() if len(pp)>1 else pp[0].strip()
        if is_house(cli) and not inc_house: continue
        ativ=str(r[1] or ''); ap=ativ.split(SEP)
        peca=limpa_titulo(ap[-1] if len(ap)>1 else ap[0])
        m=re.match(r'\s*\[([^\]]+)\]\s*(.*)',ativ)
        tipo=TIPO_MAP.get(m.group(1).strip().upper(),'Outros') if m else 'Outros'
        fmt=(m.group(2).split(SEP)[0].strip().upper() if m and m.group(2).strip() else '(SEM FORMATO)')
        fmt=fmt.split('\xa0')[0].strip() or '(SEM FORMATO)'
        das=[d.strip() for d in (str(r[8]) if r[8] else '(sem)').split(', ') if d.strip() not in EXCLUIR and d.strip() not in SAIDA]
        if not das: continue
        prazo=_prazo(r[5])
        out[(str(r[0]).strip(),ativ.strip())]={'cliente':cli,'peca':peca,'tipo':tipo,'fmt':fmt,'prio':(str(r[2]) or '').strip(),'das':das,'prazo':prazo,'status':ns(r[9] if r[9] else '')}
    return out

agora=load_pauta(f_pauta); agora_full=load_pauta(f_pauta, inc_house=True)
PAUTA=[v for v in agora.values() if any(d not in tv for d in v['das'])]
REDACAO=[v for v in agora.values() if any(d in TEAM_RED for d in v['das'])]
VIDEO=[v for v in agora.values() if any(d in TEAM_VID for d in v['das'])]
TRAF=[{'cliente':v['cliente'],'peca':v['peca'],'fmt':v['fmt'],'prazo':v['prazo'],'status':v['status']} for v in agora.values() if 'TRÁFEGO' in v['das']]
PAUTAH=[v for v in agora_full.values() if (is_house(v['cliente']) or any(d in HOUSE for d in v['das'])) and any(d not in tv for d in v['das'])]

# ---------- SNAPSHOT (movimento/novas) ----------
# chave normalizada: aplica a limpeza também em snapshots antigos (títulos 'sujos')
chave=lambda p:(str(p['cliente']).strip()+'|'+limpa_titulo(p['peca']))[:120]
snap={}
if SNAP.exists():
    s=json.loads(SNAP.read_text(encoding="utf-8"))
    if s.get("dia")==hoje: snap=s
if not snap:
    snap={"dia":hoje,"itens":[{**v} for v in PAUTA]}
    SNAP.write_text(json.dumps(snap,ensure_ascii=False),encoding="utf-8")
# arquivo histórico: guarda a foto da manhã de CADA dia (permite comparar D-1 × D)
HIST=HERE/"snapshots_hist.json"
try:
    hist=json.loads(HIST.read_text(encoding="utf-8")) if HIST.exists() else {}
except Exception: hist={}
if hoje not in hist:
    hist[hoje]={"total":len(snap["itens"]),"itens":[{'cliente':v['cliente'],'peca':v['peca'],'das':v['das'],'fmt':v.get('fmt'),'prazo':v.get('prazo')} for v in snap["itens"]]}
    for k in sorted(hist)[:-30]: hist.pop(k,None)   # mantém os últimos 30 dias
    HIST.write_text(json.dumps(hist,ensure_ascii=False),encoding="utf-8")
ch_agora={chave(p) for p in PAUTA}; ch_manha={chave(p) for p in snap["itens"]}
byk={chave(p):p for p in PAUTA}
MOVP=[{'cliente':v['cliente'],'peca':v['peca'],'tipo':v['tipo'],'prazo':v.get('prazo'),'das':v['das'],
       'entregue':chave(v) not in ch_agora,'status0':v['status'],
       'status1':byk.get(chave(v),{}).get('status')} for v in snap["itens"]]
NOVAS=[{'cliente':v['cliente'],'peca':v['peca'],'tipo':v['tipo'],'das':v['das'],'status0':v['status']} for p in PAUTA for v in [p] if chave(v) not in ch_manha]
# comparativo com o dia anterior (D-1): o que sobrou de lá e virou carga de hoje
dias_hist=sorted(d for d in hist if d<hoje)
D1={}
if dias_hist:
    d1=dias_hist[-1]; prev=hist[d1]
    ch_prev={chave(i) for i in prev['itens']}
    rollover=[i for i in prev['itens'] if chave(i) in ch_agora]
    D1={"dia":d1,"total":prev['total'],"entregues":prev['total']-len(rollover),
        "rollover":len(rollover),"itens":rollover[:80]}


# ---------- ATIVIDADES REALIZADAS ----------
def dia_iso(s):
    s=str(s); return f'{s[6:10]}-{s[3:5]}-{s[0:2]}' if len(s)>=10 and s[2]=='/' else ''
def dur_min(s):
    try:
        p=[int(x) for x in str(s).split(':')]; return round(p[0]*60+p[1]+p[2]/60,1) if len(p)==3 else 0
    except: return 0
# --------- classificação da etapa do workflow ---------
# ADMINISTRATIVA: aprovação, briefing, relatório, impulsionamento, planejamento, arquivamento, encerramento.
#   Não é entrega de peça — entra em "tempo administrativo", nunca em "peças entregues".
# PRODUTIVA: o que de fato produz a peça — ARTE na criação/vídeo e TEXTO na redação (e variantes).
# RETRABALHO: refação e ajuste — é execução, mas não é peça nova.
_ADM=re.compile(r'APROVA|BRIEFING|RELATORIO|IMPULSION|CONFERIR|PLANEJAMENTO|ARQUIVAMENTO|ENCERRAMENTO|ENECERRAMENTO|^NOVO$|CONFIGURACAO', re.I)
_REW=re.compile(r'REFACAO|AJUSTE|ALTERACAO', re.I)
_PROD=re.compile(r'ARTE|TEXTO|REDACAO|ROTEIRO|EDICAO|EDICAO|ANIMACAO|CARTELA|LETTERING|FINALIZACAO|DIAGRAMACAO|CRIACAO|REVISAO|EXECUCAO', re.I)
def _semac(x):
    return unicodedata.normalize('NFKD',str(x)).encode('ascii','ignore').decode().upper().strip()
def classe_etapa(e):
    u=_semac(e)
    if _ADM.search(u): return 'adm'
    if _REW.search(u): return 'rew'
    if _PROD.search(u): return 'prod'
    return 'out'

def net(e):
    e=str(e).strip(); u=e.upper()
    return {'APROVAÇÃO INTERNA':'Aprovação interna','REFAÇÃO INTERNA':'Refação interna','AJUSTE CLIENTE':'Ajuste cliente','AJUSTE INTERNO':'Ajuste interno','ARTE':'Arte','EDIÇÃO':'Edição','REVISÃO':'Revisão'}.get(u, e.capitalize() if e.isupper() else e)

ACT=[]; AGG=[]; MES={}; PROD={}; TEMPOS={}; TEMPOS_INT={}; TEMPOS_RED={}; TEMPOS_VID={}; das_idx=[]
AVULSO={"mes":hoje[:7],"total":0,"totalMes":0,"clientes":[],"pecas":[]}
RETRAB={"mes":hoje[:7],"pctGeral":0,"clientes":[]}; GARGALO={"mes":hoje[:7],"etapas":[]}; EFIC={"mes":hoje[:7],"real":0,"esperado":0,"pct":0,"formatos":[]}
LEAKAGE={"mes":hoje[:7],"jobs":0,"horas":0,"porDA":[],"porCliente":[]}; REWCOLAB=[]; REWDET=[]
SUP=['Diana Savi Mondo Fogaça','Gabriel Santos da Silva']  # supervisores (produção reduzida no controle de capacidade)
recs=[]; df_all=pd.DataFrame()
if VIA_API:
    # ---- atividades vindas da API do iClips ----
    _a=json.load(open(API_ATIV,encoding='utf-8'))
    df_all=pd.DataFrame([{
        'Colaborador':x.get('da',''),
        'Número do Projeto':str(x.get('idJobPeca') or ''),
        'Título do projeto':x.get('projeto',''),
        'Cliente':x.get('cliente',''),
        'Nome da peça':x.get('fmt',''),
        'Título da Peça':x.get('peca',''),
        'Etapa do Workflow':x.get('etapa',''),
        'Início':'', 'Fim':f"{x.get('dia','')[8:10]}/{x.get('dia','')[5:7]}/{x.get('dia','')[0:4]} {x.get('fim','00:00')}:00",
        'Duração':f"{int(x.get('min',0)//60):02d}:{int(x.get('min',0)%60):02d}:00",
        'Tempo estimado':'00:00:00','Custo':'',
        '_refacao':bool(x.get('refacao')), '_dep':x.get('dep','')
    } for x in _a])
    main_das={d for v in agora.values() for d in v['das'] if d not in EXCLUIR and d not in SAIDA and d not in HOUSE}
    main_das|={k for k in SAIDA if k not in EXCLUIR}
    df=df_all[df_all['Colaborador'].isin(main_das)] if len(df_all) else df_all
    recs=[]
    if len(df):
        for col,fim,cli,proj,nome,titp,etapa,dur,ref in zip(
                df['Colaborador'].astype(str), df['Fim'].astype(str), df['Cliente'].astype(str),
                df['Número do Projeto'].astype(str), df['Nome da peça'].astype(str),
                df['Título da Peça'].astype(str), df['Etapa do Workflow'].astype(str),
                df['Duração'], df['_refacao']):
            m=re.match(r'(\d{2})/(\d{2})/(\d{4})\s+(\d{2}:\d{2})',fim)
            if not m: continue
            dia=f'{m.group(3)}-{m.group(2)}-{m.group(1)}'
            _col=col.strip()
            if _col in EXCLUIR: continue
            _s=SAIDA.get(_col)
            if _s and dia>=_s: continue
            mt=re.match(r'\s*\[([^\]]+)\]\s*(.*)',nome)
            fmt=(mt.group(2).strip().upper() if mt and mt.group(2).strip() else nome.strip().upper() or '(SEM FORMATO)')
            tit=limpa_titulo(titp if titp and titp!='nan' else nome)[:90]
            # a API já diz se é refação: usa o campo em vez de adivinhar pelo nome da etapa
            cl='rew' if ref else classe_etapa(etapa)
            recs.append((_col,dia,m.group(4),cli.strip(),proj,tit,fmt,net(etapa),dur_min(dur),cl))
    # HISTÓRICO: os .xls antigos continuam valendo para os dias que a API ainda não trouxe.
    # Assim a migração para a API não apaga o histórico do ano.
    _dias_api={x[1] for x in recs}
    if f_ativ:
        _dfs=[]
        for _p in f_ativ:
            try:
                _pk=_p.parent/(_p.name+'.pkl')
                if _pk.exists() and _pk.stat().st_mtime>=_p.stat().st_mtime: _dfs.append(pd.read_pickle(_pk))
                else:
                    _d=pd.read_html(_p, encoding='iso-8859-1')[0]
                    try: _d.to_pickle(_pk)
                    except Exception: pass
                    _dfs.append(_d)
            except Exception: pass
        if _dfs:
            _h=pd.concat(_dfs,ignore_index=True).drop_duplicates(subset=['Colaborador','Fim','Nome da peça'])
            _h=_h[_h['Colaborador'].isin(main_das)].dropna(subset=['Fim'])
            _nhist=0
            for col,fim,cli,proj,nome,titp,etapa,dur in zip(
                    _h['Colaborador'].astype(str),_h['Fim'].astype(str),_h['Cliente'].astype(str),
                    _h['Número do Projeto'].astype(str),_h['Nome da peça'].astype(str),
                    _h['Título da Peça'].astype(str),_h['Etapa do Workflow'].astype(str),_h['Duração']):
                m=re.match(r'(\d{2})/(\d{2})/(\d{4})\s+(\d{2}:\d{2})',fim)
                if not m: continue
                dia=f'{m.group(3)}-{m.group(2)}-{m.group(1)}'
                if dia in _dias_api: continue          # dia já veio da API: não duplica
                _col=col.strip()
                if _col in EXCLUIR: continue
                _s2=SAIDA.get(_col)
                if _s2 and dia>=_s2: continue
                mt=re.match(r'\s*\[([^\]]+)\]\s*(.*)',nome)
                fmt=(mt.group(2).strip().upper() if mt and mt.group(2).strip() else nome.strip().upper() or '(SEM FORMATO)')
                tit=limpa_titulo(titp if titp and titp!='nan' else nome)[:90]
                recs.append((_col,dia,m.group(4),cli.strip(),proj,tit,fmt,net(etapa),dur_min(dur),classe_etapa(etapa)))
                _nhist+=1
            if _nhist: print(f"   histórico dos .xls: +{_nhist} apontamentos de dias não cobertos pela API")
elif f_ativ:
    dfs=[]
    for p in f_ativ:
        try:
            pk=p.parent/(p.name+'.pkl')   # cache rápido: lê o .xls uma vez, reusa o pickle depois
            if pk.exists() and pk.stat().st_mtime>=p.stat().st_mtime:
                dfs.append(pd.read_pickle(pk))
            else:
                d=pd.read_html(p, encoding='iso-8859-1')[0]
                try: d.to_pickle(pk)
                except Exception: pass
                dfs.append(d)
        except Exception: pass
    df_all=pd.concat(dfs, ignore_index=True).drop_duplicates(subset=['Colaborador','Fim','Nome da peça']) if dfs else pd.DataFrame()
    # quem está na pauta de hoje + quem saiu do time (para o histórico do ano não sumir com a pessoa)
    main_das={d for v in agora.values() for d in v['das'] if d not in EXCLUIR and d not in SAIDA and d not in HOUSE}
    main_das|={k for k in SAIDA if k not in EXCLUIR}
    df=df_all[df_all['Colaborador'].isin(main_das)].dropna(subset=['Fim']) if len(df_all) else df_all
    recs=[]
    if len(df):
        _z=zip(df['Colaborador'].astype(str),df['Fim'].astype(str),df['Cliente'].astype(str),
               df['Número do Projeto'].astype(str),df['Nome da peça'].astype(str),
               df['Título da Peça'].astype(str),df['Etapa do Workflow'].astype(str),df['Duração'])
        for col,fim,cli,proj,nome,titp,etapa,dur in _z:
            m=re.match(r'(\d{2})/(\d{2})/(\d{4})\s+(\d{2}:\d{2})',fim)
            if not m: continue
            dia=f'{m.group(3)}-{m.group(2)}-{m.group(1)}'
            mt=re.match(r'\s*\[([^\]]+)\]\s*(.*)',nome)
            fmt=(mt.group(2).strip().upper() if mt and mt.group(2).strip() else nome.strip().upper() or '(SEM FORMATO)')
            tit=limpa_titulo(titp if titp and titp!='nan' else nome)[:90]
            _col=col.strip()
            if _col in EXCLUIR: continue            # removida do dashboard em qualquer data
            _s=SAIDA.get(_col)
            if _s and dia>=_s: continue             # saiu do time: histórico anterior fica, daqui pra frente não conta
            recs.append((_col,dia,m.group(4),cli.strip(),proj,tit,fmt,net(etapa),dur_min(dur),classe_etapa(etapa)))

# ---- processamento comum: vale para a API e para os exports manuais ----
if recs:
    dias_all=sorted({x[1] for x in recs}); recent=set(dias_all[-20:])
    ACT=[{'da':a,'dia':d,'fim':f,'cliente':c,'pid':p,'peca':t,'fmt':fm,'etapa':e,'dur':du,'cl':cl} for a,d,f,c,p,t,fm,e,du,cl in recs if d in recent]
    das_idx=sorted({x[0] for x in recs}); IDX={d:i for i,d in enumerate(das_idx)}
    # p  = peças ENTREGUES  -> só conta quem passou por etapa produtiva (Arte / Texto / Edição...) ou retrabalho
    # pa = peças que tiveram SÓ etapa administrativa (aprovação, briefing, relatório) — informativo, fora da entrega
    # ma = horas gastas em etapas administrativas
    agg=defaultdict(lambda:[set(),0,0,0,set(),0]); mes=defaultdict(lambda:[set(),0])
    for a,d,f,c,p,t,fm,e,du,cl in recs:
        g=agg[(d,IDX[a])]; key=p+'|'+t
        g[1]+=1; g[2]+=du
        if cl=='adm':
            g[4].add(key); g[5]+=du
        else:
            g[0].add(key)
            if cl=='rew': g[3]+=1
        if cl!='adm':
            mk=(IDX[a],d[:7]); mes[mk][0].add(key); mes[mk][1]+=du
    AGG=[{'dia':d,'d':i,'p':len(v[0]),'a':v[1],'m':round(v[2]),'r':v[3],
          'pa':len(v[4]-v[0]),'ma':round(v[5])} for (d,i),v in sorted(agg.items())]
    MES=defaultdict(dict)
    for (i,mm),v in mes.items(): MES[i][mm]=[len(v[0]),round(v[1])]
    # tempos por formato (mediana da etapa Arte) — vetorizado p/ performance
    POOL_INT={'Arthur Andre Souza Santos','Alexandre Ramos dos Santos','Gustavo da Silva Rangel Aragonez','Guilherme Gonçalves Ferreira','Thayssa Polyanna Souza Lira','Lucas Reis Chaves'}
    tg=defaultdict(list); tgi=defaultdict(list)
    dfa=df_all[df_all['Etapa do Workflow'].astype(str).str.strip().str.upper().isin(['ARTE','CRIAÇÃO DE ARTE'])].copy()
    if len(dfa):
        dfa['du']=dfa['Duração'].map(dur_min)
        dfa=dfa[(dfa['du']>1)&(dfa['du']<=2880)]
        dfa['f']=dfa['Nome da peça'].fillna('').astype(str).str.extract(r'^\s*\[[^\]]+\]\s*(.+)$')[0].str.strip().str.upper()
        dfa=dfa[dfa['f'].notna()]
        for f_,du in zip(dfa['f'],dfa['du']): tg[f_].append(du)
        di=dfa[dfa['Colaborador'].astype(str).str.strip().isin(POOL_INT)]
        for f_,du in zip(di['f'],di['du']): tgi[f_].append(du)
    TEMPOS={f:[round(statistics.median(v)),len(v)] for f,v in tg.items() if len(v)>=8}
    TEMPOS_INT={f:[round(statistics.median(v)),len(v)] for f,v in tgi.items() if len(v)>=5}
    # tempo TOTAL por peça (soma das etapas) por formato — equipes de redação e vídeo
    def _fmt_series(s):
        return s.fillna('').astype(str).str.replace(r'^\s*\[[^\]]+\]\s*','',regex=True).str.split('\xa0').str[0].str.strip().str.upper().str[:40].replace('','(SEM)')
    def _tt(members):
        dd=df_all[df_all['Colaborador'].isin(members)].copy()
        if not len(dd): return {}
        dd['min']=dd['Duração'].map(dur_min)
        dd['pid']=dd['Número do Projeto'].astype(str)+'|'+dd['Título da Peça'].astype(str)
        dd['fmt']=_fmt_series(dd['Nome da peça'])
        per=dd.groupby(['pid','fmt'],sort=False)['min'].sum().reset_index()
        per=per[(per['min']>0)&(per['min']<=1200)]
        grp=defaultdict(list)
        for fmt,mn in zip(per['fmt'],per['min']): grp[fmt].append(mn)
        out={f:[round(statistics.median(v)),len(v)] for f,v in grp.items() if len(v)>=5}
        allv=[x for v in grp.values() for x in v]
        if allv: out['__def__']=[round(statistics.median(allv)),len(allv)]
        return out
    TEMPOS_RED=_tt(TEAM_RED); TEMPOS_VID=_tt(TEAM_VID)

    # ===== PROD — tempo em play x produtividade por pessoa (semana e mês) =====
    # Regra: só etapa produtiva/retrabalho conta como peça. Etapa administrativa entra só como tempo.
    # "esperado" = mediana histórica do formato NO GRUPO da pessoa (redator não é comparado com editor).
    def _grp_de(n):
        v=ALOC.get(n)              # no robô ALOC é dict de tuplas: (grupo, foco)
        return v[0] if v else ''
    def _esp(da,fmt):
        g=_grp_de(da)
        tb = TEMPOS_RED if g=='Redação' else TEMPOS_VID if g=='Vídeo' else TEMPOS_INT if g in ('Interno','Mídia Off') else None
        if tb and fmt in tb: return tb[fmt][0]
        if tb and '__def__' in tb: return tb['__def__'][0]
        return (TEMPOS.get(fmt) or [60])[0]
    def _iso(d):
        y,m_,dd_=int(d[:4]),int(d[5:7]),int(d[8:10])
        iy,iw,_=datetime.date(y,m_,dd_).isocalendar()
        return f"{iy}-S{iw:02d}"
    _pp=defaultdict(lambda:{'m':0,'esp':0,'rew':0,'adm':0,'pc':set(),'dias':set()})
    for a,d,f,c,p,t,fm,e,du,cl in recs:
        for per in (_iso(d), d[:7]):
            x=_pp[(a,per)]
            x['dias'].add(d)
            if cl=='adm': x['adm']+=du; continue
            x['m']+=du
            if cl=='rew': x['rew']+=du
            k=p+'|'+t
            if k not in x['pc']:
                x['pc'].add(k); x['esp']+=_esp(a,fm)
    PROD={}
    for (a,per),x in _pp.items():
        if not x['pc']: continue
        PROD.setdefault(per,[]).append({
            'da':a,'g':_grp_de(a),'d':len(x['dias']),'m':round(x['m']),'p':len(x['pc']),
            'esp':round(x['esp']),'rew':round(x['rew']),'adm':round(x['adm'])})
    for per in PROD: PROD[per].sort(key=lambda r:-r['p'])
    PROD={k:v for k,v in sorted(PROD.items()) if k>='2026-01'}
    # peças avulsas (fora de escopo) do mês corrente
    mes=hoje[:7]
    dm=df_all[df_all['Fim'].astype(str).str.match(r'\d{2}/\d{2}/\d{4}',na=False)].copy()
    dm['mm']=dm['Fim'].astype(str).str.slice(3,10).str.replace(r'(\d{2})/(\d{4})',r'\2-\1',regex=True)
    dm=dm[dm['mm']==mes]
    if len(dm):
        dm['pid']=dm['Número do Projeto'].astype(str)+'|'+dm['Título da Peça'].astype(str)
        pcs=dm.drop_duplicates('pid')
        pcs=pcs[~pcs['Cliente'].map(is_house)]   # análises excluem clientes house (só aparecem na aba deles)
        isav=pcs['Nome da peça'].astype(str).str.contains('avuls',case=False,na=False)|pcs['Título da Peça'].astype(str).str.contains('avuls',case=False,na=False)
        av=pcs[isav]
        from collections import Counter as _C
        tc=_C(pcs['Cliente'].astype(str)); rc=_C(av['Cliente'].astype(str))
        clientes=[{"cliente":c,"n":n,"tot":tc[c],"pct":round(100*n/max(1,tc[c]))} for c,n in rc.most_common()]
        pecas=[]
        for _,r in av.iterrows():
            mt=re.match(r'\s*\[[^\]]+\]\s*(.*)',str(r['Nome da peça'] or ''))
            fmt=((mt.group(1).split('\xa0')[0].strip()[:40]) if mt and mt.group(1).strip() else str(r['Nome da peça'] or '').strip()[:40])
            tit=limpa_titulo(r['Título da Peça'])[:90]
            dd=re.match(r'(\d{2})/(\d{2})/(\d{4})',str(r['Fim']))
            dia=f"{dd.group(3)}-{dd.group(2)}-{dd.group(1)}" if dd else ""
            pecas.append({"cliente":str(r['Cliente'] or ''),"peca":tit,"fmt":fmt,"dia":dia})
        pecas.sort(key=lambda x:(-1,x['cliente'],x['dia']))
        pecas.sort(key=lambda x:x['cliente'])
        AVULSO={"mes":mes,"total":len(av),"totalMes":len(pcs),"clientes":clientes,"pecas":pecas}
        # ---- indicadores analíticos (mês) ----
        dmc=dm.copy(); dmc['min']=dmc['Duração'].map(dur_min)
        RETc=re.compile(r'refa|ajuste|reprov',re.I)
        dmc['ret']=dmc['Etapa do Workflow'].astype(str).map(lambda e:bool(RETc.search(e)))
        dmi=dmc[~dmc['Cliente'].map(is_house)].copy()   # base das análises: só clientes NÃO-house
        # 1) RETRABALHO por cliente (sem house)
        rrows=[]
        for cli,d in dmi.groupby('Cliente'):
            pcsn=d['pid'].nunique()
            if pcsn<10: continue
            pr=d[d['ret']]['pid'].nunique(); hr=round(d[d['ret']]['min'].sum()/60,1)
            rrows.append({"cliente":str(cli),"pecas":int(pcsn),"comRetr":int(pr),"pct":round(100*pr/max(1,pcsn)),"horas":hr})
        rrows.sort(key=lambda x:-x['comRetr'])
        totp=dmi['pid'].nunique(); totpr=dmi[dmi['ret']]['pid'].nunique()
        RETRAB={"mes":mes,"pctGeral":round(100*totpr/max(1,totp)),"clientes":rrows}
        # 2) GARGALO por etapa
        def _net(e):
            e=str(e).strip().upper()
            for k in ['ARTE','TEXTO','EDIÇÃO','EDICAO','ANIMAÇÃO','ANIMACAO','REVISÃO','REVISAO','ROTEIRO','APROVAÇÃO','APROVACAO','REFAÇÃO','REFACAO','AJUSTE','ARQUIVAMENTO','CLAQUETE','CARTELA','STORYBOARD']:
                if k in e: return {'EDICAO':'EDIÇÃO','ANIMACAO':'ANIMAÇÃO','REVISAO':'REVISÃO','APROVACAO':'APROVAÇÃO','REFACAO':'REFAÇÃO'}.get(k,k)
            return 'OUTROS'
        dmi['et']=dmi['Etapa do Workflow'].map(_net)
        gg=dmi.groupby('et').agg(ativ=('min','size'),minutos=('min','sum'),pecas=('pid','nunique'))
        tmin_g=gg['minutos'].sum() or 1
        GARGALO={"mes":mes,"totalMin":round(tmin_g),"etapas":[{"etapa":i,"min":round(r['minutos']),"pct":round(100*r['minutos']/tmin_g),"pecas":int(r['pecas']),"ativ":int(r['ativ'])} for i,r in gg.sort_values('minutos',ascending=False).iterrows()]}
        # 3) EFICIÊNCIA esperado (mediana histórica total/formato) × real (mês)
        allp=df_all.copy(); allp['min']=allp['Duração'].map(dur_min)
        allp['pid']=allp['Número do Projeto'].astype(str)+'|'+allp['Título da Peça'].astype(str); allp['fmt']=_fmt_series(allp['Nome da peça'])
        histp=allp.groupby(['pid','fmt'])['min'].sum().reset_index()
        esp={};
        for f,v in histp.groupby('fmt')['min'].apply(list).items():
            if len(v)>=20 and 0<statistics.median(v)<=1200: esp[f]=round(statistics.median(v))
        dmi['fmt']=_fmt_series(dmi['Nome da peça'])
        realp=dmi.groupby(['pid','fmt'])['min'].sum().reset_index()
        frows=[]; tR=0; tE=0
        for f,d in realp.groupby('fmt'):
            if f not in esp: continue
            vals=[x for x in d['min'] if 0<x<=1200]
            if len(vals)<10: continue   # só formatos com amostra sólida (evita distorção)
            rm=round(statistics.median(vals))
            frows.append({"fmt":f,"esperado":esp[f],"real":rm,"n":len(vals),"desvio":round(100*(rm-esp[f])/esp[f])})
            tR+=sum(vals); tE+=esp[f]*len(vals)
        frows.sort(key=lambda x:-x['n'])
        EFIC={"mes":mes,"real":round(tR/60),"esperado":round(tE/60),"pct":round(100*tR/max(1,tE)),"formatos":frows}
        # ---- LEAKAGE: nosso time (não-house) cobrindo cliente house ----
        nosso_nh={n for n,v in ALOC.items() if v[0]!='House'}
        lk=dmc[dmc['Cliente'].map(is_house) & dmc['Colaborador'].isin(nosso_nh)]
        porDA=[];
        for da,d in lk.groupby('Colaborador'):
            porDA.append({"nome":da,"grupo":ALOC.get(da,('',''))[0],"jobs":int(d['pid'].nunique()),"horas":round(d['min'].sum()/60,1)})
        porDA.sort(key=lambda x:-x['jobs'])
        porCli=[{"cliente":str(c),"jobs":int(d['pid'].nunique())} for c,d in lk.groupby('Cliente')]
        porCli.sort(key=lambda x:-x['jobs'])
        LEAKAGE={"mes":mes,"jobs":int(lk['pid'].nunique()),"horas":round(lk['min'].sum()/60,1),"porDA":porDA,"porCliente":porCli}
        # ---- RETRABALHO por colaborador (nosso quadro) + detalhe p/ drill-down ----
        roster=set(ALOC.keys())
        rc=[]
        for col,d in dmi[dmi['Colaborador'].isin(roster)].groupby('Colaborador'):
            p=d['pid'].nunique()
            if p<8: continue
            pr=d[d['ret']]['pid'].nunique()
            rc.append({"nome":col,"grupo":ALOC.get(col,('',''))[0],"pecas":int(p),"comRetr":int(pr),"pct":round(100*pr/max(1,p)),"horas":round(d[d['ret']]['min'].sum()/60,1)})
        rc.sort(key=lambda x:-x['comRetr']); REWCOLAB=rc
        det=[]
        for _,r in dmi[dmi['ret']].iterrows():
            tit=limpa_titulo(r['Título da Peça'])[:70]
            dd=re.match(r'(\d{2})/(\d{2})/(\d{4})',str(r['Fim']))
            det.append({"colab":str(r['Colaborador']),"cliente":str(r['Cliente'] or ''),"peca":tit,"etapa":str(r['Etapa do Workflow']).strip()[:26],"min":round(r['min']),"dia":(f"{dd.group(3)}-{dd.group(2)}-{dd.group(1)}" if dd else "")})
        REWDET=det

# nomes curtos
all_names=sorted(set(das_idx) | {d for p in PAUTA for d in p['das']} | tv | set(ALOC.keys()))
SHORT={}; seen=set()
for n in all_names:
    w=n.split(); s=(w[0]+' '+w[-1]) if len(w)>1 else n
    if s in seen and len(w)>2: s=f'{w[0]} {w[1][0]}. {w[-1]}'
    seen.add(s); SHORT[n]=s

aloc_js={k:{'g':v[0],'f':v[1]} for k,v in ALOC.items()}
ent=sum(1 for p in MOVP if p['entregue'])
SUB=(f"Farol da gestão produtiva · criação {len(PAUTA)} · house {len(PAUTAH)} · redação {len(REDACAO)} · "
     f"vídeo {len(VIDEO)} · tráfego {len(TRAF)} · {ent} concluídos desde a manhã · "
     f"atualizado {datetime.datetime.now():%d/%m %Hh%M}")

html=TPL
for ph,val in [('__PAUTA__',PAUTA),('__MOVP__',MOVP),('__NOVAS__',NOVAS),('__ACT__',ACT),('__DAS__',das_idx),
               ('__AGG__',AGG),('__MES__',MES),('__PROD__',PROD),('__AUSENTES__',AUSENTES.get(hoje,[])),('__TRAF__',TRAF),('__RED__',REDACAO),('__VID__',VIDEO),
               ('__TEAMRED__',list(TEAM_RED)),('__TEAMVID__',list(TEAM_VID)),('__ALOC__',aloc_js),('__SHORT__',SHORT),
               ('__TEMPOS__',TEMPOS),('__TEMPOSINT__',TEMPOS_INT),('__PAUTAH__',PAUTAH),
               ('__TEMPOSRED__',TEMPOS_RED),('__TEMPOSVID__',TEMPOS_VID),('__SUP__',SUP),('__AVULSO__',AVULSO),
               ('__RETRAB__',RETRAB),('__GARGALO__',GARGALO),('__EFIC__',EFIC),
               ('__LEAKAGE__',LEAKAGE),('__REWCOLAB__',REWCOLAB),('__REWDET__',REWDET),('__D1__',D1)]:
    html=html.replace(ph, json.dumps(val, ensure_ascii=False))
html=html.replace('__REF__',hoje).replace('__BUILDTS__',datetime.datetime.now().strftime('%Y-%m-%dT%H:%M'))
# logo oficial da Vanguarda (símbolo V) embutido em base64
_lg=HERE/"logo_v.b64"
if _lg.exists(): html=html.replace('__LOGOB64__', _lg.read_text().strip())
resto=[m for m in re.findall(r'__[A-Z]+__', html)]
if resto: sys.exit(f"ERRO: placeholders não preenchidos: {sorted(set(resto))} (o template está desatualizado?)")
# injeta a biblioteca de PowerPoint (após o guard, para não confundir o validador)
_pl=HERE/"pptxgen_lib.js"
html=html.replace('PPTXLIBSLOT', _pl.read_text(encoding='utf-8') if _pl.exists() else '')
(HERE/"index.html").write_text(html, encoding="utf-8")
print(f"[{datetime.datetime.now():%d/%m %H:%M}] fonte: {f_pauta.name} | criação {len(PAUTA)} · house {len(PAUTAH)} · red {len(REDACAO)} · vid {len(VIDEO)} · traf {len(TRAF)} · vencem hoje {sum(1 for p in PAUTA if p['prazo']==hoje)} · concluídos {ent}")

# ---------- DEPLOY NETLIFY ----------
tok=cfg.get("netlify_token","")
if not tok or tok.startswith("COLE_"):
    print("index.html gerado (sem deploy: preencha netlify_token)."); sys.exit(0)
h={"Authorization":f"Bearer {tok}"}
sites=requests.get("https://api.netlify.com/api/v1/sites", headers=h, timeout=60).json()
site=next((s for s in sites if cfg["netlify_site"] in (s.get("name",""),(s.get("custom_domain") or ""))), None)
if not site: sys.exit(f"Site '{cfg['netlify_site']}' não encontrado no Netlify.")
content=html.encode("utf-8"); sha=hashlib.sha1(content).hexdigest()
dep=requests.post(f"https://api.netlify.com/api/v1/sites/{site['id']}/deploys", headers=h, json={"files":{"/index.html":sha}}, timeout=90).json()
if sha in dep.get("required",[]):
    requests.put(f"https://api.netlify.com/api/v1/deploys/{dep['id']}/files/index.html",
                 headers={**h,"Content-Type":"application/octet-stream"}, data=content, timeout=90).raise_for_status()
print(f"Publicado no Netlify ✓ ({site.get('ssl_url') or site.get('url')})")
