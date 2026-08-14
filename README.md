# Farol da Criação

Painel de pauta, capacidade e entregas da produção criativa da **Vanguarda Martech**.

Lê os dados do iClips, calcula capacidade por profissional e gera um `index.html`
autocontido — funciona offline, dá para mandar por WhatsApp e publicar no Netlify.

---

## O que o painel mostra

| Aba | Conteúdo |
|---|---|
| **Demandas Internas** | pauta do núcleo de criação e mídia off, capacidade por DA, redistribuição |
| **Movimento do dia** | o que entrou, saiu e foi remanejado desde a manhã |
| **Redação** | pauta, tempo por formato e painel de capacidade de entrega |
| **Produção de Vídeo** | pauta e tempos do time de audiovisual |
| **Tráfego** | demanda por dia do mês |
| **Houses · Dedicados** | pauta dos DAs alocados em cliente |
| **Nova Era & Pátio** | pauta dos clientes house |
| **Play × Produtividade** | horas apontadas × peças entregues × ritmo, por semana e mês |
| **Peças avulsas** | demandas fora do escopo no mês |

---

## Como rodar

### Modo API (recomendado)

```bash
py src/farol_api.py 2026-01-01   # carga histórica, uma vez só
py src/farol_pasta.py            # gera o index.html
```

No dia a dia: **clique duplo em `atualizar.bat`**.

### Modo manual (reserva)

Coloque em `dados/` os dois exports do iClips e rode:

```bash
py src/farol_pasta.py
```

- `lista_peças*.xlsx` — Gestão de Atividades (a pauta)
- `relatorio*.xls` — Atividade realizada por funcionário

---

## Configuração

Antes do primeiro uso, crie dois arquivos a partir dos exemplos:

```bash
copy api_key.exemplo.txt api_key.txt
copy farol_config.exemplo.json farol_config.json
```

- **`api_key.txt`** — chave do iClips (Avatar → Chave de API)
- **`farol_config.json`** — token do Netlify, se for publicar

> Os dois estão no `.gitignore`. **Nunca** faça commit deles.

---

## Estrutura

```
src/                       código
  farol_pasta.py           gera o dashboard
  farol_api.py             extrai do iClips (somente leitura)
  analise_pauta_classe.py  pauta por classe de cliente (AA/A/B/C)
  farol_template_deploy.html  template do painel
dados/                     exports do iClips (fora do Git)
docs/                      manuais e memória da operação
farol_constantes.json      equipe, alocação, exclusões, ausências
clientes_classes.json      classificação de clientes (AA/A/B/C)
```

---

## Regras de negócio que valem lembrar

- **Peça entregue hoje estava na pauta de hoje** — entra em "jobs no dia" e sai do backlog.
- **Etapa administrativa não é entrega.** Aprovação, briefing, relatório e arquivamento
  contam como tempo, nunca como peça. Vale Arte na criação e Texto na redação.
- **Capacidade é o tempo até o fim do expediente.** Não desconta horas apontadas,
  porque o timesheet do iClips tem timers esquecidos.
- **Cada quadro da visão geral espelha uma aba** — o número é sempre conferível.
- **Quem sai do time** entra em `SAIDA` com a data: o histórico anterior é preservado.

---

## Segurança

Este repositório **não contém** chaves, tokens nem dados pessoais.
Os relatórios do iClips trazem nome, produtividade e horas de cerca de 40 pessoas —
por isso a pasta `dados/` está fora do versionamento.

**Mantenha o repositório privado.**
