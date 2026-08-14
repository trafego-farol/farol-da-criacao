# Memória da operação — Farol da Criação · Vanguarda Martech
_Atualizado em 12/08/2026_

## Classificação de clientes (novo — 12/08/2026)

Fonte: iClips → **Clientes atendimento** (3 páginas), 113 clientes ativos.
Arquivo de dados: **`clientes_classes.json`** · Análise: **`analise_pauta_classe.py`**

| Classe | Clientes | Peso | Significado na alocação |
|---|---|---|---|
| **AA** | 37 | 4 | Prioridade máxima — atender primeiro |
| **A** | 37 | 3 | Prioridade alta |
| **B** | 23 | 2 | Prioridade média |
| **C** | 5 | 1 | Prioridade menor |
| **—** | 11 | 0 | Ativo sem classe e carteira 0 — revisar cadastro antes de priorizar |

**Regra de alocação:** AA → A → B → C. Os "sem classe" ficam para checagem cadastral
ou tratamento pontual, não entram na fila normal de prioridade.

### Apelidos — o iClips grava nomes diferentes do cadastro
Mapeados em `clientes_classes.json → APELIDOS`. Os que existiam até agora:

| Nome na pauta (iClips) | Cadastro oficial |
|---|---|
| Grupo Braga | BRAGA VEÍCULOS VAREJO |
| IAA \| CAA | IAA - INDÚSTRIA DO ALUMÍNIO |
| MILLENNIUM CENTER | MILLENNUM SHOPPING |
| OPEN MALL AMAZON | AMAZON OPEN MALL |
| Grupo Nova Era | NOVA ERA |
| CAMPANHA THEREZINHA RUIZ 2026 | THEREZINHA RUIZ |
| PMZ \| ESCOLA DE MECÂNICOS | PMZ ESCOLA DE MECÂNICOS |
| CDLM | CDL MANAUS |
| VANGUARDA MIDIA DIGITAL | VANGUARDA COMUNICAÇÃO (interno) |

> Sempre que aparecer cliente novo com classe `?` na análise, acrescentar o apelido aqui
> e em `clientes_classes.json`.

### Grupos econômicos (juntar na mesma janela de atendimento)
FOGÁS (10 empresas) · BRAGA (5) · PMZ (3) · PNEU FORTE (2) · NOVA ERA (5) ·
PÁTIO GOURMET (2) · OLÁ / REALIZE (9) · TROPICAL (4) · CAA / IAA (3) ·
SÃO PEDRO (3) · RIO / SPE (2) · INTERNO VANGUARDA (2)

### Leitura da pauta de 12/08/2026 (296 peças)
- **41% da pauta é cliente AA** (122 peças em 19 clientes) — PMZ 20, Therezinha Ruiz 16,
  Santo Remédio 15, Yamaha 14, Unipar 10
- A 25% · B 21% · sem classe 12%
- Grupo **BRAGA aparece fatiado em 6 empresas** (20 peças somadas) — candidato natural
  a janela única de atendimento
- **37 peças (12%) são de clientes sem classe**, quase todas Nova Era e Pátio Gourmet,
  que são houses e por isso não têm carteira

## Onde isso NÃO está
Por decisão do Matheus, a classificação **não entra no dashboard**. Ela vive no banco de
dados (`clientes_classes.json`) e é usada na análise da pauta pelo script, para orientar
a distribuição do dia.

## Como usar
```
py analise_pauta_classe.py                    # usa dados\lista_pecas.xlsx
py analise_pauta_classe.py caminho\outra.xlsx
```
Devolve: peças por classe, ordem de atendimento por cliente, grupos econômicos na pauta
e carga por profissional ponderada pelo peso do cliente.

## Removidos do dashboard

| Pessoa | Quando | Motivo |
|---|---|---|
| Victor Lima dos Santos | — | fora do escopo do painel |
| Luana Rocha Souza | — | fora do escopo do painel |
| Gabriela Lee Guedes de Souza | saída em 04/08/2026 | histórico de jan a 03/08 preservado |
| Bianca Rodrigues Raposo da Camara | 14/08/2026 | removida a pedido; não tinha apontamentos |
