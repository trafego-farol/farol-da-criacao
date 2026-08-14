# Publicar no GitHub — 4 passos

O repositório **já está pronto e commitado**. Falta só conectar à sua conta.

---

## 1. Instale o Git (se ainda não tiver)

https://git-scm.com/download/win — instale com as opções padrão.

Para conferir, abra o Prompt e digite:

```
git --version
```

---

## 2. Crie o repositório no GitHub

1. Entre em https://github.com/new
2. **Repository name:** `farol-da-criacao`
3. **Marque `Private`** — o projeto tem configuração de equipe e clientes
4. **NÃO marque** "Add a README file", nem `.gitignore`, nem licença.
   Precisa estar **vazio**, senão o envio dá conflito.
5. Clique em **Create repository**
6. Copie a URL que aparece, algo como
   `https://github.com/seu-usuario/farol-da-criacao.git`

---

## 3. Copie esta pasta para o seu computador

Coloque onde preferir. Ela já tem o histórico do Git dentro (pasta `.git`).

---

## 4. Publique

Clique duas vezes em **`publicar_github.bat`**, cole a URL quando ele pedir.

Na primeira vez o Git abre o navegador para você entrar na sua conta —
é o jeito seguro, você não precisa digitar senha em lugar nenhum.

Se preferir pelo Prompt:

```
git remote add origin https://github.com/seu-usuario/farol-da-criacao.git
git push -u origin main
```

---

## Depois, no dia a dia

```
git add -A
git commit -m "descreva o que mudou"
git push
```

---

## O que NÃO vai para o GitHub

O `.gitignore` já bloqueia:

- `api_key.txt` — chave do iClips
- `farol_config.json` — token do Netlify
- `dados/` — os exports com nome, horas e produtividade de ~40 pessoas
- `index.html` e os `Farol_*.html` — saída gerada, não código

O `publicar_github.bat` **confere isso antes de enviar** e cancela se achar
algo sensível.

---

## Se der erro

| Mensagem | O que fazer |
|---|---|
| `failed to push some refs` | o repositório foi criado com README. Crie um novo, vazio. |
| `Authentication failed` | conclua o login na janela do navegador |
| `repository not found` | confira a URL e se o repositório é seu |
| `git não é reconhecido` | Git não instalado — volte ao passo 1 |
