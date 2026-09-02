/* =========================================================================
 * Agenda DC — leitura de planilha Excel e renderização da agenda.
 *
 * A planilha padrão fica em data/agenda.xlsx. Edite-a (ou substitua-a) e
 * a agenda atualiza automaticamente ao recarregar / republicar o site.
 * Também é possível carregar uma planilha local via "Carregar planilha".
 * ========================================================================= */

'use strict';

const PLANILHA_PADRAO = 'data/agenda.xlsx';

const ORDEM_DIAS = [
  'segunda-feira', 'terça-feira', 'quarta-feira',
  'quinta-feira', 'sexta-feira', 'sábado', 'domingo',
];

// Aliases aceitos para os nomes de colunas (tudo comparado sem acento/caixa).
const COLUNAS = {
  dia:         ['dia', 'dia da semana', 'diadasemana', 'weekday'],
  data:        ['data', 'date'],
  inicio:      ['inicio', 'hora inicio', 'horainicio', 'inicio', 'começo', 'comeco', 'start'],
  fim:         ['fim', 'hora fim', 'horafim', 'termino', 'end'],
  atividade:   ['atividade', 'titulo', 'título', 'evento', 'descricao', 'descrição', 'assunto', 'title'],
  local:       ['local', 'sala', 'lugar', 'location', 'place'],
  responsavel: ['responsavel', 'responsável', 'palestrante', 'facilitador', 'owner'],
  categoria:   ['categoria', 'tipo', 'category'],
  observacoes: ['observacoes', 'observações', 'obs', 'notas', 'notes'],
};

const CORES_CATEGORIA = {
  'reuniao':     'var(--cat-reuniao)',
  'reunião':     'var(--cat-reuniao)',
  'treinamento': 'var(--cat-treinamento)',
  'evento':      'var(--cat-evento)',
  'pausa':       'var(--cat-pausa)',
  'almoço':      'var(--cat-pausa)',
  'almoco':      'var(--cat-pausa)',
  'coffee':      'var(--cat-pausa)',
  'tarefa':      'var(--cat-tarefa)',
};

let eventos = []; // modelo normalizado em memória

/* ------------------------------- Utilidades ------------------------------ */

function semAcento(s) {
  return String(s ?? '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .trim().toLowerCase();
}

function corDaCategoria(categoria) {
  return CORES_CATEGORIA[semAcento(categoria)] || 'var(--cat-outro)';
}

function capitalizar(s) {
  s = String(s ?? '').trim();
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : s;
}

/** Converte valores de hora (string "9:00", número serial do Excel, Date) em "HH:MM". */
function normalizarHora(valor) {
  if (valor === null || valor === undefined || valor === '') return '';
  if (valor instanceof Date) {
    return String(valor.getHours()).padStart(2, '0') + ':' +
           String(valor.getMinutes()).padStart(2, '0');
  }
  if (typeof valor === 'number') {
    // Fração do dia usada pelo Excel para horas (0.5 = 12:00).
    const totalMin = Math.round((valor % 1) * 24 * 60);
    const h = Math.floor(totalMin / 60);
    const m = totalMin % 60;
    return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
  }
  const txt = String(valor).trim();
  const m = txt.match(/(\d{1,2})[:h.]?(\d{2})?/);
  if (m) {
    const h = m[1].padStart(2, '0');
    const min = (m[2] || '00').padStart(2, '0');
    return `${h}:${min}`;
  }
  return txt;
}

function normalizarData(valor) {
  if (!valor) return '';
  let d = valor;
  if (typeof valor === 'number') {
    // Serial do Excel (dias desde 1899-12-30).
    d = new Date(Math.round((valor - 25569) * 86400 * 1000));
  } else if (!(valor instanceof Date)) {
    const parsed = new Date(valor);
    if (isNaN(parsed)) return String(valor);
    d = parsed;
  }
  return d.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit' });
}

/** Mapeia os cabeçalhos reais da planilha para as chaves do nosso modelo. */
function mapearColunas(cabecalhos) {
  const mapa = {};
  cabecalhos.forEach((cab) => {
    const chave = semAcento(cab).replace(/\s+/g, ' ');
    for (const [campo, aliases] of Object.entries(COLUNAS)) {
      if (aliases.some((a) => semAcento(a) === chave)) {
        mapa[campo] = cab;
        break;
      }
    }
  });
  return mapa;
}

/* ------------------------- Leitura da planilha --------------------------- */

function processarWorkbook(workbook) {
  const nomePlanilha = workbook.SheetNames[0];
  const sheet = workbook.Sheets[nomePlanilha];
  const linhas = XLSX.utils.sheet_to_json(sheet, { defval: '', raw: true });

  if (!linhas.length) {
    eventos = [];
    return;
  }

  const mapa = mapearColunas(Object.keys(linhas[0]));

  eventos = linhas
    .map((linha) => ({
      dia:         String(linha[mapa.dia] ?? '').trim(),
      data:        normalizarData(linha[mapa.data]),
      inicio:      normalizarHora(linha[mapa.inicio]),
      fim:         normalizarHora(linha[mapa.fim]),
      atividade:   String(linha[mapa.atividade] ?? '').trim(),
      local:       String(linha[mapa.local] ?? '').trim(),
      responsavel: String(linha[mapa.responsavel] ?? '').trim(),
      categoria:   String(linha[mapa.categoria] ?? '').trim(),
      observacoes: String(linha[mapa.observacoes] ?? '').trim(),
    }))
    .filter((e) => e.atividade || e.dia); // ignora linhas totalmente vazias
}

async function carregarPlanilhaPadrao() {
  mostrarStatus('Carregando planilha…');
  try {
    const resp = await fetch(PLANILHA_PADRAO + '?t=' + Date.now());
    if (!resp.ok) throw new Error('HTTP ' + resp.status);
    const buffer = await resp.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
    processarWorkbook(workbook);
    renderizarTudo();
    mostrarStatus(`Planilha carregada · ${eventos.length} atividades.`, false, 2500);
  } catch (err) {
    console.error(err);
    eventos = [];
    renderizarTudo();
    mostrarStatus(
      'Não foi possível carregar data/agenda.xlsx. Use "Carregar planilha" para enviar um arquivo.',
      true
    );
  }
}

function carregarArquivoLocal(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const workbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
      processarWorkbook(workbook);
      renderizarTudo();
      mostrarStatus(`"${file.name}" carregada · ${eventos.length} atividades.`, false, 3000);
    } catch (err) {
      console.error(err);
      mostrarStatus('Erro ao ler o arquivo. Verifique se é um Excel válido.', true);
    }
  };
  reader.readAsArrayBuffer(file);
}

/* ---------------------------- Renderização ------------------------------- */

function chaveDia(evento) {
  // Usa a data quando disponível para diferenciar dias iguais em semanas distintas.
  return evento.dia || evento.data || 'Sem dia';
}

function ordenarPorHora(a, b) {
  return (a.inicio || '99:99').localeCompare(b.inicio || '99:99');
}

function diaDeHoje() {
  return ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira',
          'quinta-feira', 'sexta-feira', 'sábado'][new Date().getDay()];
}

function agruparPorDia() {
  const grupos = new Map();
  eventos.forEach((ev) => {
    const k = chaveDia(ev);
    if (!grupos.has(k)) grupos.set(k, { dia: ev.dia, data: ev.data, itens: [] });
    grupos.get(k).itens.push(ev);
  });
  // Ordena os dias pela ordem natural da semana, depois alfabético.
  return [...grupos.values()].sort((a, b) => {
    const ia = ORDEM_DIAS.indexOf(semAcento(a.dia));
    const ib = ORDEM_DIAS.indexOf(semAcento(b.dia));
    if (ia !== -1 && ib !== -1) return ia - ib;
    if (ia !== -1) return -1;
    if (ib !== -1) return 1;
    return String(a.dia).localeCompare(String(b.dia));
  });
}

function renderizarTudo() {
  atualizarFiltroCategorias();
  renderizarSemana();
  atualizarSeletorDia();
  renderizarDia();
}

/* --------- Visão semanal --------- */

function renderizarSemana() {
  const container = document.getElementById('grade-semana');
  const filtro = document.getElementById('filtro-categoria-semana').value;
  const grupos = agruparPorDia();

  if (!grupos.length) {
    container.innerHTML = estadoVazio('Nenhuma atividade encontrada na planilha.');
    return;
  }

  const hoje = diaDeHoje();
  container.innerHTML = grupos.map((grupo) => {
    const itens = grupo.itens
      .filter((ev) => !filtro || ev.categoria === filtro)
      .sort(ordenarPorHora);

    const ehHoje = semAcento(grupo.dia) === hoje;

    const cards = itens.length
      ? itens.map((ev) => `
          <div class="evento-card" style="border-left-color:${corDaCategoria(ev.categoria)}">
            ${ev.inicio ? `<div class="hora">${ev.inicio}${ev.fim ? '–' + ev.fim : ''}</div>` : ''}
            <div class="titulo">${escapar(ev.atividade)}</div>
            ${ev.local ? `<div class="meta">📍 ${escapar(ev.local)}</div>` : ''}
          </div>`).join('')
      : '<div class="dia-vazio">Sem atividades</div>';

    return `
      <div class="dia-coluna ${ehHoje ? 'hoje' : ''}">
        <div class="dia-header">
          <div class="nome">${escapar(capitalizar(grupo.dia) || 'Sem dia')}${ehHoje ? '<span class="badge-hoje">Hoje</span>' : ''}</div>
          ${grupo.data ? `<div class="data">${escapar(grupo.data)}</div>` : ''}
        </div>
        <div class="dia-eventos">${cards}</div>
      </div>`;
  }).join('');

  document.getElementById('semana-info').textContent =
    `${eventos.length} atividades · ${grupos.length} dias.`;
}

/* --------- Visão diária --------- */

function atualizarSeletorDia() {
  const seletor = document.getElementById('seletor-dia');
  const grupos = agruparPorDia();
  const anterior = seletor.value;

  seletor.innerHTML = grupos.map((g) => {
    const rotulo = capitalizar(g.dia || 'Sem dia') + (g.data ? ` (${g.data})` : '');
    return `<option value="${escapar(chaveDia(g.itens[0]))}">${escapar(rotulo)}</option>`;
  }).join('');

  // Preserva a seleção ou escolhe o dia de hoje quando existir.
  const hoje = diaDeHoje();
  const grupoHoje = grupos.find((g) => semAcento(g.dia) === hoje);
  if (anterior && [...seletor.options].some((o) => o.value === anterior)) {
    seletor.value = anterior;
  } else if (grupoHoje) {
    seletor.value = chaveDia(grupoHoje.itens[0]);
  }
}

function renderizarDia() {
  const container = document.getElementById('timeline-dia');
  const seletor = document.getElementById('seletor-dia');
  const chave = seletor.value;

  const itens = eventos
    .filter((ev) => chaveDia(ev) === chave)
    .sort(ordenarPorHora);

  if (!itens.length) {
    container.innerHTML = estadoVazio('Selecione um dia com atividades.');
    document.getElementById('dia-info').textContent = 'Linha do tempo com horários.';
    return;
  }

  document.getElementById('dia-info').textContent =
    `${itens.length} atividades programadas.`;

  container.innerHTML = itens.map((ev) => {
    const cor = corDaCategoria(ev.categoria);
    const chips = [
      ev.local ? `<span class="chip">📍 ${escapar(ev.local)}</span>` : '',
      ev.responsavel ? `<span class="chip">👤 ${escapar(ev.responsavel)}</span>` : '',
      ev.categoria ? `<span class="chip categoria" style="background:${cor}">${escapar(ev.categoria)}</span>` : '',
    ].join('');

    return `
      <div class="timeline-item">
        <div class="timeline-hora">
          ${ev.inicio || '—'}
          ${ev.fim ? `<span class="fim">${ev.fim}</span>` : ''}
        </div>
        <div class="timeline-card" style="border-left-color:${cor}">
          <p class="titulo">${escapar(ev.atividade)}</p>
          <div class="chips">${chips}</div>
          ${ev.observacoes ? `<p class="obs">${escapar(ev.observacoes)}</p>` : ''}
        </div>
      </div>`;
  }).join('');
}

/* --------- Filtros e auxiliares de UI --------- */

function atualizarFiltroCategorias() {
  const select = document.getElementById('filtro-categoria-semana');
  const atual = select.value;
  const categorias = [...new Set(eventos.map((e) => e.categoria).filter(Boolean))].sort();
  select.innerHTML = '<option value="">Todas</option>' +
    categorias.map((c) => `<option value="${escapar(c)}">${escapar(c)}</option>`).join('');
  if ([...select.options].some((o) => o.value === atual)) select.value = atual;
}

function estadoVazio(msg) {
  return `<div class="empty-state">
    <div class="big">🗓️</div>
    <h3>Nada por aqui</h3>
    <p>${escapar(msg)}</p>
  </div>`;
}

function escapar(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

let statusTimer;
function mostrarStatus(msg, erro = false, autoHide = 0) {
  const bar = document.getElementById('status-bar');
  bar.textContent = msg;
  bar.classList.toggle('error', erro);
  bar.hidden = false;
  clearTimeout(statusTimer);
  if (autoHide) statusTimer = setTimeout(() => (bar.hidden = true), autoHide);
}

function trocarView(view) {
  document.querySelectorAll('.tab').forEach((t) => {
    const ativo = t.dataset.view === view;
    t.classList.toggle('is-active', ativo);
    t.setAttribute('aria-selected', ativo);
  });
  document.getElementById('view-semana').hidden = view !== 'semana';
  document.getElementById('view-dia').hidden = view !== 'dia';
}

/* ------------------------ Planilha modelo (CSV) -------------------------- */

function baixarModeloCSV(e) {
  e.preventDefault();
  const cabecalho = 'Dia,Data,Início,Fim,Atividade,Local,Responsável,Categoria,Observações';
  const exemplo = 'Segunda-feira,08/09,09:00,10:00,Reunião de abertura,Sala 1,Ana,Reunião,Trazer relatório';
  const csv = '﻿' + cabecalho + '\n' + exemplo + '\n';
  const url = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' }));
  const a = document.createElement('a');
  a.href = url;
  a.download = 'agenda-modelo.csv';
  a.click();
  URL.revokeObjectURL(url);
}

/* ------------------------------- Eventos --------------------------------- */

function iniciar() {
  document.querySelectorAll('.tab').forEach((t) =>
    t.addEventListener('click', () => trocarView(t.dataset.view)));

  document.getElementById('file-input').addEventListener('change', (e) => {
    if (e.target.files[0]) carregarArquivoLocal(e.target.files[0]);
  });
  document.getElementById('btn-reload').addEventListener('click', carregarPlanilhaPadrao);
  document.getElementById('filtro-categoria-semana').addEventListener('change', renderizarSemana);
  document.getElementById('seletor-dia').addEventListener('change', renderizarDia);
  document.getElementById('link-modelo').addEventListener('click', baixarModeloCSV);

  carregarPlanilhaPadrao();
}

document.addEventListener('DOMContentLoaded', iniciar);
