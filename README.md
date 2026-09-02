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

## Estrutura

```
agenda-DC/
├── index.html                     # página principal (abas Semana / Dia)
├── css/styles.css                 # estilos
├── js/app.js                      # leitura do Excel + renderização
├── data/agenda.xlsx               # dados da agenda (edite este arquivo)
├── scripts/gerar_planilha_exemplo.py
└── .github/workflows/deploy.yml   # deploy automático no GitHub Pages
```

## Publicando no GitHub Pages

No repositório: **Settings → Pages → Build and deployment → Source: GitHub Actions**.
A partir daí, cada push na branch `main` publica a versão mais recente.
