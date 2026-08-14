# Farol da Criação — Documentação Técnica

**Vanguarda Martech** · Painel de gestão da produção criativa
Última atualização deste documento: 28/07/2026

---

## 1. O que é

Dashboard operacional que mostra, em tempo quase real: a **pauta do dia** por profissional, a **capacidade** de entrega da equipe, o **movimento do dia** (o que entrou/saiu/foi concluído) e o **histórico de produção** do ano.

Fonte de dados: **iClips** (sistema de gestão da agência), via 2 relatórios exportados em Excel.

**Stack:** HTML único autossuficiente (dados embutidos) + Chart.js 4.5 via CDN + gerador em Python 3.
**Sem backend, sem banco de dados, sem build step.** O arquivo final abre offline e pode ser enviado por e-mail/WhatsApp.

---

## 2. Arquitetura

```
┌─────────────────────┐
│ iClips (2 exports)  │  ← passo manual: exportar para a pasta /dados
│ • lista_peças*.xlsx │     (Gestão de Atividades = PAUTA)
│ • relatorio*.xls    │     (Atividades Realizadas = EXECUÇÃO)
└──────────┬──────────┘
           ↓
┌─────────────────────────────────────────┐
│ farol_pasta.py  (gerador / "robô")      │
│ 1. lê o arquivo MAIS RECENTE de cada    │
│    tipo na pasta /dados                 │
│ 2. calcula todos os indicadores         │
│ 3. injeta os dados como JSON no         │
│    template HTML (placeholders __X__)   │
│ 4. grava index.html                     │
│ 5. publica no Netlify (API)             │
└──────────┬──────────────────────────────┘
           ↓
┌─────────────────────────────────┐
│ index.html (1,4 MB)             │
│ dados embutidos como const JS   │
│ publicado em                    │
│ vanguarda-farol.netlify.app     │
└─────────────────────────────────┘
```

Agendamento atual: **Agendador de Tarefas do Windows**, a cada 30 min, chamando `Farol_Pasta.bat`.

### Arquivos do projeto

| Arquivo | Função |
|---|---|
| `farol_pasta.py` | Gerador. Lê os Excel, calcula, preenche o template, publica |
| `farol_template_deploy.html` | Template com placeholders `__PAUTA__`, `__AGG__`, `__BUILDTS__`, `__LOGOB64__`… |
| `farol_constantes.json` | Alocação dos ~45 profissionais (grupo + cliente), times de Redação/Vídeo, exclusões |
| `farol_config.json` | Token e nome do site no Netlify **(não versionar — contém segredo)** |
| `logo_v.b64` | Logo Vanguarda (símbolo V) em base64, injetado no HTML |
| `snapshot_pauta.json` | Foto da pauta na 1ª execução do dia (base do "movimento do dia") |
| `snapshots_hist.json` | Histórico das fotos diárias, últimos 30 dias (base do comparativo D-1) |
| `dados/*.pkl` | Cache dos relatórios (pandas pickle). Acelera o build de ~40s para ~8s |
| `Farol_Pasta.bat` | Atalho: `py farol_pasta.py >> farol_log.txt 2>&1` |

**Dependências Python:** `openpyxl`, `pandas`, `lxml`, `requests`.

---

## 3. Modelo de dados

### Entrada A — Gestão de Atividades (a PAUTA)
Colunas: `Projeto | Atividade | Prioridade | Data início | Hora | Data fim | Hora | Data de Conclusão | Responsáveis | Status da Peça`

- **Cliente** = última parte de `Projeto` após o separador `\xa0\xa0·\xa0\xa0`
- **Peça / formato** = `Atividade`, no padrão `[TIPO] FORMATO  ·  descrição`
- **Prazo** = `Data fim` · **Responsáveis** = lista separada por `, ` · **Status** = `Status da Peça`

### Entrada B — Atividades Realizadas (a EXECUÇÃO)
Colunas: `Colaborador | Número do Projeto | Título do projeto | Cliente | Nome da peça | Título da Peça | Etapa do Workflow | Início | Fim | Duração | Tempo estimado | Custo`

- Uma linha **por etapa executada**. `Fim` = conclusão da etapa. `Duração` = HH:MM:SS
- Chave da peça = `Número do Projeto` + `Título da Peça`
- ⚠️ Os `.xls` do iClips são **HTML disfarçado** — leia com `pandas.read_html(..., encoding='iso-8859-1')`, não com `read_excel`

### Estruturas injetadas no HTML (constantes JS)

| Constante | Conteúdo |
|---|---|
| `PAUTA` | Pauta atual (exclui clientes house). `{cliente, peca, tipo, fmt, prio, das[], prazo, status}` |
| `PAUTA_HOUSE` | Pauta dos clientes house (Nova Era, Pátio Gourmet) |
| `PAUTA_INT` / `PAUTA_DED` / `PAUTA_CORE` | Derivadas: interna, dedicados, núcleo (interna − dedicados) |
| `REDACAO` / `VIDEO` / `TRAF` | Pauta por especialidade |
| `MOVP` | Foto da manhã com flag `entregue` (saiu da pauta) |
| `NOVAS` | Peças que entraram após a foto da manhã |
| `AGG` | Agregado por dia × profissional: `{dia, d, p:peças, a:atividades, m:minutos, r:retrabalho}` |
| `ACT` | Atividades detalhadas (últimos 20 dias) |
| `MES` | Produção mensal por profissional |
| `TEMPOS` / `TEMPOS_INT` / `TEMPOS_RED` / `TEMPOS_VID` | Mediana de minutos por formato (histórico) |
| `AVULSO` / `RETRAB` / `GARGALO` / `EFICD` / `LEAKAGE` | Indicadores analíticos do mês |
| `D1` | Comparativo com o dia anterior (rollover) |
| `ALOC` / `SHORT` / `TEAM_RED` / `TEAM_VID` / `SUPSET` | Metadados de pessoas |
| `REF` / `BUILDTS` | Data de referência e timestamp do build |

---

## 4. Regras de negócio (importante manter)

**Frentes** — pela tag `[TIPO]` no início da atividade:
`[SOCIAL MEDIA]`→Social Media · `[INBOUND]`→Inbound · `[OFF]`→Off · `[DEV]`→Dev/On · `[MKT DIGITAL]`→Mídia Paga · `[SCIENCE SEO]`→SEO · sem tag→Outros

**Grupos de profissionais** (em `farol_constantes.json`): House · Interno (núcleo digital) · Dedicado · Mídia Off · Vídeo · Redação · Freela.

**Separação por aba (regra dura):** Redação, Vídeo e Tráfego têm abas próprias e **não aparecem** na tabela do setor de Arte. Clientes house aparecem **somente** na aba deles — as análises (retrabalho, avulsas, gargalo, eficiência) usam apenas clientes **não-house**.

**Capacidade:** jornada de **8h a 100%** para todos (`EFIC = 1.0`, `SUPF = 1.0` — gestores não têm redução). Limite de peças/dia por DA configurável (padrão 8).

**Ocupação = horas em aberto ÷ jornada (8h).** ⚠️ Não dividir pelo tempo restante do dia: isso infla o indicador ao longo da tarde (chegava a 270%). O risco de não terminar é exibido **à parte**, como aviso "não cabe hoje", quando `horas_em_aberto > tempo_restante_até_fim_do_expediente`.

**Capacidade restante** = minutos entre agora e o fim do expediente, usando **hora local** (`localISO()`, nunca UTC), recalculada a cada 3 min via `setInterval`.

**Entregue / movimento:** uma peça conta como concluída quando **sai da pauta** entre a foto da manhã e a extração atual. As entregas por profissional vêm do relatório de atividades (`AGG` do dia).

**"Peças no dia" por profissional** = em aberto **+** já entregues hoje (a carga real do dia). **Backlog** = apenas o que continua em aberto.

**Escopo × extraescopo:** peça com "**Avulso**" no nome → extraescopo. A tag `{ESCOPO}` existe mas é preenchida de forma inconsistente no iClips — **não confiar nela** como critério principal.

**Limpeza de título (ETL):** `limpa_titulo()` remove a tag `[TIPO]`, `{datas}`, `(ESCOPO)` e separadores soltos. ⚠️ A **chave de comparação** entre snapshots aplica essa mesma limpeza — sem isso, snapshots antigos (títulos "sujos") geram falsas entregas.

**Tempos por formato:** mediana do histórico, só formatos com amostra suficiente (`n ≥ 10` na eficiência; `n ≥ 8` nos tempos gerais). Formatos raros usam o padrão da equipe.

---

## 5. Front-end

**Layout:** shell em CSS Grid — `aside#side` (menu lateral, 262px, colapsável para 66px) + `.mainarea`.

**Navegação:** sem abas horizontais. Menu lateral agrupado em Visão geral · Especialidades · Carteiras · Análises, com **badges de contagem** (quantas vencem hoje por frente) e **breadcrumb**. Estado do menu persistido em `localStorage` (`sideMini`).

**Páginas** (`.page`, uma visível por vez): `p1` Demandas Internas · `p2` Movimento do dia · `p4` Redação · `p5` Vídeo · `p6` Tráfego · `p8` Houses·Dedicados · `p7` Nova Era & Pátio · `p3` Entregas & histórico · `p9` Peças avulsas.

**Lazy init:** os gráficos são criados **só quando a aba é aberta** (`if(!inited[p]) INIT[p]()`).
⚠️ `showTab()` já chama `INIT` — **não chamar `INIT.p1()` manualmente** depois, ou o Chart.js quebra com *"Canvas is already in use"*.

**Seções ocultáveis:** cada `.sec[data-sec]` tem botão "Ocultar"; o estado vai para `localStorage` (`farolHid3`). Seções de **consulta** (governança, plano, tempos, tendência, mensal…) já começam **recolhidas** — lista em `HID_PADRAO`.

### Organização dos controles (padrão a manter)
- **Parâmetros fixos** (limite/dia, jornada, meta, fim do expediente) → **barra lateral**, bloco `.side-par`
- **Seletores que mudam a página** (dia da pauta, dia em análise) → **dentro da página que afetam**, em `.ctrl`, visíveis e rotulados
- **Cabeçalho** = apenas marca, selo de atualização e exportar

### Identidade visual (cores oficiais, extraídas do logo `.ai`)

| Token | Hex | Uso |
|---|---|---|
| `--navy` | `#1D1E1C` | Grafite da marca: cabeçalho, sidebar, títulos |
| `--red` | `#E11121` | **Somente atenção**: alertas, atrasos, sobrecarga, item ativo do menu |
| `--green` | `#2E7D32` | Positivo, entregue |
| `--orange` | `#ED7D31` | Atenção intermediária |
| `--grey` | `#F4F4F4` | Fundo |
| `--border` | `#e4e5e3` | Bordas |

**Regra de cor:** vermelho é escasso por decisão de design (feedback de UX: "o vermelho chama muita [atenção]"). Cabeçalho é grafite com um fio vermelho de 3px; item ativo do menu usa fundo translúcido + indicador vermelho de 3px, não bloco vermelho.

**Tipografia:** Inter (fallback Segoe UI / system-ui). Números grandes com `letter-spacing` negativo.

**Responsivo:** breakpoints em **900px** (menu vira faixa horizontal rolável no topo, header empilha, grids 2 colunas, tabelas com scroll-x e 1ª coluna congelada) e **520px** (grids 1 coluna). Tabelas: `.scroll` com `overflow-x:auto`, `min-width:660px` e `position:sticky` na primeira coluna.

**Acessibilidade:** ícones do menu com `aria-hidden="true"`, botões com `aria-label`, `<nav aria-label>`.

**Impressão:** `@media print` isola a página 1 e esconde navegação/controles.

---

## 6. Como rodar / manter

```bash
pip install openpyxl pandas lxml requests

# 1. exportar os 2 relatórios do iClips para ./dados
#    (nome DEVE começar com "lista" e "relatorio")
# 2. rodar
py farol_pasta.py
# saída: index.html + publish no Netlify (se houver token no config)
```

**Publicação:** API do Netlify — `GET /api/v1/sites` para achar o site, `POST /sites/{id}/deploys` com o SHA1 do arquivo, `PUT` do conteúdo.

### Pegadinhas conhecidas

- Nomes dos arquivos em `/dados` precisam começar com **`lista`** (pauta) e **`relatorio`** (execução), senão o robô não os encontra
- O relatório grande (~22 MB) leva ~14s na primeira leitura; depois usa o `.pkl` em cache
- Vários relatórios de execução podem coexistir na pasta: são concatenados e deduplicados por `Colaborador + Fim + Nome da peça`
- Para manter o histórico do ano, deixe um relatório longo (01/01→hoje) fixo na pasta e adicione os curtos do dia
- O snapshot da manhã define o "movimento do dia" — a 1ª execução do dia deve acontecer cedo

---

## 7. Backlog técnico (prioridade sugerida)

1. **Coleta automática.** Hoje a exportação dos 2 relatórios é manual. A API pública do iClips (Plano PRO, header `platform: {token}`, endpoint `content-data`) foi testada: serve para entregas/retrabalho/histórico, mas **não reproduz a pauta do dia com fidelidade** (é orientada a BI, uma linha por etapa). Existe um `farol_api_puller.py` pronto, em stand-by, aguardando validação.
2. **SLA de entrega no prazo** — exige um export do iClips com *prazo + data de conclusão real* por peça (a Gestão de Atividades atual só traz as peças em aberto).
3. **Entregas por profissional com precisão** — o relatório de "concluídas" do iClips vem com a coluna *Responsáveis* vazia; hoje resolvemos cruzando pelo número do projeto com o relatório de atividades.
4. **Cornerstone (planejado × executado)** — cruzar VJob (planejamento) com iClips (execução). Depende de um *de-para* de clientes e formatos entre os dois sistemas.
5. **n8n** — 2 fluxos prontos: vigia (alerta se o painel parar de atualizar) e resumo diário. Leem o HTML publicado, rodam na nuvem.

---

## 8. Contexto de decisões (por que está assim)

- **Dados embutidos no HTML, e não API/banco:** requisito de negócio — o arquivo precisa abrir offline, ser enviável por WhatsApp e não depender de infraestrutura. 1,4 MB carrega em ~1s. Não há problema de performance a resolver aqui.
- **Robô de pasta, e não integração via API:** a API do iClips foi testada primeiro e falhou em reproduzir a pauta; a versão por export manual dá 100% de fidelidade.
- **9 páginas, e não uma tela única:** houve uma tentativa de reorganizar em "home executiva + 4 seções" que foi revertida por decisão do usuário — a operação preferiu o acesso direto. A simplificação foi feita **dentro** de cada página (remoção de KPIs duplicados, seções de consulta recolhidas por padrão).
