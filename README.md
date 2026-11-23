# 📅 Agenda DC

Agenda web que se alimenta de **planilhas do Excel**. Você mantém a programação
numa planilha e a agenda é montada automaticamente, com duas visões:

- **📆 Semana** — programação geral da semana, um cartão por dia.
- **🕒 Dia** — programação detalhada de um dia, em linha do tempo com horários.

Não há backend: é um site estático (HTML + CSS + JavaScript) que lê o Excel
diretamente no navegador usando [SheetJS](https://sheetjs.com).

---

## Como funciona

1. Os dados ficam em **`data/agenda.xlsx`**.
2. Ao abrir o site, a planilha é lida e a agenda é renderizada.
3. Para atualizar a agenda, **edite `data/agenda.xlsx`** e faça `git push`.
   O GitHub Pages republica o site automaticamente (workflow em
   `.github/workflows/deploy.yml`).
4. Também é possível **carregar uma planilha do seu computador** pelo botão
   *"Carregar planilha"* — útil para pré-visualizar sem fazer commit.

---

## Formato da planilha

Uma aba com as colunas abaixo (a ordem não importa; acentos são opcionais):

| Coluna        | Obrigatória | Exemplo             |
|---------------|:-----------:|---------------------|
| `Dia`         | sim         | Segunda-feira       |
| `Data`        | não         | 08/09               |
| `Início`      | recomendada | 09:00               |
| `Fim`         | não         | 10:00               |
| `Atividade`   | sim         | Reunião de abertura |
| `Local`       | não         | Sala 1              |
| `Responsável` | não         | Ana                 |
| `Categoria`   | não         | Reunião             |
| `Observações` | não         | Trazer relatório    |

### Categorias com cor própria

`Reunião`, `Treinamento`, `Evento`, `Pausa` (ou `Almoço`) e `Tarefa`.
Qualquer outra categoria usa a cor neutra.

Um modelo em CSV pode ser baixado pelo link no rodapé do site.

---

## Rodando localmente

Por causa da leitura de arquivo via `fetch`, use um servidor local (abrir o
`index.html` direto pelo `file://` pode bloquear o carregamento da planilha):

```bash
# a partir da raiz do projeto
python3 -m http.server 8000
# depois acesse http://localhost:8000
```

## Regenerar a planilha de exemplo

```bash
pip install openpyxl
python3 scripts/gerar_planilha_exemplo.py
```

---

## 🎂 Aniversariantes do dia no WhatsApp

Todo dia às **08:00 (Brasília)** o workflow
`.github/workflows/aniversariantes.yml` lê a planilha de aniversariantes e
envia para **um grupo do WhatsApp** quem faz aniversário no dia, com nome,
cidade, cargo e telefone:

```
📅 *Segunda-feira, 23/11/2026*
🎉 *Aniversariantes de hoje*

🎂 *Maria Exemplo*
📍 Campinas — Vereadora
📞 (19) 99999-0000
💬 Mensagem privada: https://wa.me/5519999990000
```

O link 💬 abre direto uma conversa privada com o aniversariante, para quem
quiser mandar os parabéns pessoalmente. Nada é enviado automaticamente a ele.
Telefones sem DDI são tratados como brasileiros (+55).

Nos dias sem aniversariante, nada é enviado. Para mandar um aviso nesses dias,
crie a variável `MENSAGEM_SEM_ANIVERSARIANTES`.

### ⚠️ Os dados pessoais não ficam neste repositório

O repositório é público, então a planilha com nomes e telefones **não pode ser
commitada** (o `.gitignore` bloqueia `data/aniversariantes*`). Os logs do
Actions também são públicos, por isso o script não imprime nomes nem telefones
quando roda lá.

### Formato da planilha

Excel (`.xlsx`) ou CSV. A ordem das colunas não importa e os acentos são
opcionais. Veja o modelo em `data/aniversariantes.exemplo.csv`.

| Coluna       | Obrigatória | Exemplo                       |
|--------------|:-----------:|-------------------------------|
| `Nome`       | sim         | Maria Silva                   |
| `Nascimento` | sim         | 23/11, 23/11/1980 ou data do Excel |
| `Cidade`     | não         | Campinas                      |
| `Cargo`      | não         | Vereadora                     |
| `Telefone`   | não         | 19999990000                   |

Quem nasceu em 29/02 aparece em 28/02 nos anos não bissextos.

### Configuração (Settings → Secrets and variables → Actions)

**1. Onde está a planilha** (crie um destes *Secrets*):

- `ANIVERSARIANTES_FONTE`: um link para baixar a planilha. Com Google Sheets:
  *Compartilhar → Qualquer pessoa com o link (leitor)* e use
  `https://docs.google.com/spreadsheets/d/<ID>/export?format=xlsx`.
  O link fica guardado como segredo e não aparece em lugar nenhum.
- ou `ANIVERSARIANTES_CSV`: o próprio conteúdo do CSV colado no segredo.

**2. Para onde enviar** (escolha o provedor na *Variable* `WHATSAPP_PROVEDOR`:
`zapi`, que é o padrão, ou `evolution`):

| Provedor  | Secrets                                                   |
|-----------|-----------------------------------------------------------|
| Z-API     | `ZAPI_INSTANCIA`, `ZAPI_TOKEN`, `ZAPI_CLIENT_TOKEN`        |
| Evolution | `EVOLUTION_URL`, `EVOLUTION_INSTANCIA`, `EVOLUTION_APIKEY` |
| Ambos     | `WHATSAPP_GRUPO_ID`: o ID do grupo                         |

O ID do grupo é obtido no painel do provedor (na Z-API, em "listar grupos":
algo como `120363012345678901-group`; na Evolution, `120363012345678901@g.us`).
Recomendo usar um **chip dedicado** conectado ao provedor e adicionado ao grupo.

**3. Personalização opcional** (*Variables*): `MENSAGEM_TITULO`,
`MENSAGEM_RODAPE` (texto no fim da mensagem; vazio por padrão) e
`MENSAGEM_SEM_ANIVERSARIANTES`.

### Testando

- **No GitHub:** *Actions → Aniversariantes do dia → Run workflow*. Com
  "Modo teste" marcado, o workflow lê a planilha e mostra quantos
  aniversariantes encontrou, mas não envia nada. Desmarque para enviar de
  verdade. Também é possível simular outro dia.
- **Enviar um teste só para você** (antes de divulgar no grupo):
  1. Crie o secret `WHATSAPP_TESTE_NUMERO` com o seu número, ex.: `5519999990000`.
  2. Em *Run workflow*, **desmarque** "Modo teste", **marque** "Enviar só
     para mim" e, se hoje não tiver aniversariante, preencha "Simular outro
     dia" com uma data que tenha (ex.: `25/12`).
  3. A mensagem chega no seu WhatsApp com o aviso "🧪 TESTE" no topo. O grupo
     não recebe nada.
- **Rodar manualmente se o automático atrasar:** *Run workflow* com "Modo
  teste" desmarcado. Existe uma trava de **um envio por dia ao grupo**: se o
  automático começar depois (ou já tiver enviado), ele vê o registro do dia e
  não envia de novo. Para reenviar de propósito, marque "Forçar reenvio".
  Testes ("Modo teste", "Enviar só para mim" ou "Simular outro dia") não
  contam para a trava.
- **No computador:**

```bash
ANIVERSARIANTES_FONTE=data/aniversariantes.exemplo.csv python3 scripts/aniversariantes.py --data 23/11 --dry-run
```

> O agendamento só roda depois que o workflow estiver na branch `main`. Em
> repositórios públicos, o GitHub pausa agendamentos depois de 60 dias sem
> commits. Se isso acontecer, basta reativar em *Actions*.

---

## Estrutura

```
agenda-DC/
├── index.html                     # página principal (abas Semana / Dia)
├── css/styles.css                 # estilos
├── js/app.js                      # leitura do Excel + renderização
├── data/agenda.xlsx               # dados da agenda (edite este arquivo)
├── data/aniversariantes.exemplo.csv  # modelo (dados fictícios)
├── scripts/gerar_planilha_exemplo.py
├── scripts/aniversariantes.py     # envio diário dos aniversariantes
└── .github/workflows/
    ├── deploy.yml                 # deploy automático no GitHub Pages
    └── aniversariantes.yml        # agendamento diário do WhatsApp
```

## Publicando no GitHub Pages

No repositório: **Settings → Pages → Build and deployment → Source: GitHub Actions**.
A partir daí, cada push na branch `main` publica a versão mais recente.
