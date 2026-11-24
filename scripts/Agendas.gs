// Calendário Olímpico CTO – agendas públicas do Google para as famílias assinarem.
//
// Cria 3 agendas na conta que autorizar este script (deve ser a conta do CTO):
//   Anos Iniciais (1º ao 5º ano), Anos Finais (6º ao 9º ano) e Ensino Médio.
// A cada 2 horas copia para elas as datas PÚBLICAS da planilha. Prazos internos,
// reuniões e tudo marcado "Somente equipe" ficam de fora.
//
// Instalação (uma vez, no computador, logado só na conta do CTO):
//   1. Extensões → Apps Script → botão + ao lado de "Arquivos" → Script → nome "Agendas" → cole este código.
//   2. Em "Serviços" (+), adicione "Google Calendar API".
//   3. Escolha a função instalarAgendas no menu de cima e clique em Executar. Autorize.

const AGENDAS_CTO = [
  {id: 'efai', seg: 'EFAI', nome: 'Olimpíadas CTO · Anos Iniciais (1º ao 5º ano)'},
  {id: 'efaf', seg: 'EFAF', nome: 'Olimpíadas CTO · Anos Finais (6º ao 9º ano)'},
  {id: 'em',   seg: 'EM',   nome: 'Olimpíadas CTO · Ensino Médio'},
];
const SITE_CTO = 'https://ctomackenziebrasilia.github.io/';
const ETAPAS_CTO = [ // nome, coluna início, coluna fim, interna
  ['Inscrição', 'inscricao (inicio)', 'inscricao (fim)', false],
  ['Divulgação aos alunos', 'divulgacao aos alunos', null, false],
  ['Prazo interno CTO', 'prazo interno cto', null, true],
  ['1ª Fase', '1ª fase (inicio)', '1ª fase (fim)', false],
  ['2ª Fase', '2ª fase (inicio)', '2ª fase (fim)', false],
  ['Fase Final', 'fase final (inicio)', 'fase final (fim)', false],
  ['Resultado', 'resultado', null, false],
  ['Premiação', 'premiacao', null, false],
];
const TIPOS_INTERNOS_CTO = ['prazo interno', 'prazo interno cto', 'reuniao'];
const MESES_CTO = ['janeiro','fevereiro','março','abril','maio','junho','julho','agosto','setembro','outubro','novembro','dezembro'];
const JANELA_CURTA_CTO = 14; // períodos de até 14 dias viram um evento só
const LIMITE_MS_CTO = 5 * 60 * 1000; // para antes do limite de 6 min; a próxima rodada continua

// Rode UMA vez: cria as agendas, deixa públicas, agenda a sincronização e sincroniza.
function instalarAgendas() {
  AGENDAS_CTO.forEach(a => agendaCTO_(a, true));
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'sincronizarAgendas')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('sincronizarAgendas').timeBased().everyHours(2).create();
  sincronizarAgendas();
}

// Menu na planilha para sincronizar na hora.
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Calendário CTO')
    .addItem('Sincronizar agendas públicas agora', 'sincronizarAgendas')
    .addItem('Definir senhas do painel', 'definirSenhasPainel') // função do arquivo Painel
    .addToUi();
}

// Copia as datas públicas da planilha para as 3 agendas (cria, altera e apaga o necessário).
function sincronizarAgendas() {
  const inicio = Date.now();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const eventos = montarEventosCTO_();
    let feitas = 0, pendente = false;
    for (const a of AGENDAS_CTO) {
      const cal = agendaCTO_(a);
      const quero = {};
      eventos.filter(e => e.segs.indexOf(a.seg) >= 0).forEach(e => quero[e.uid] = e);
      const existentes = cal.getEvents(new Date(2020, 0, 1), new Date(2040, 0, 1));
      const vistos = {};
      for (const ev of existentes) {
        if (Date.now() - inicio > LIMITE_MS_CTO) { pendente = true; break; }
        const uid = ev.getTag('cto');
        const e = uid && quero[uid];
        if (!e || vistos[uid]) { ev.deleteEvent(); feitas++; continue; }
        vistos[uid] = true;
        const fimExcl = maisDias_(e.fim, 1);
        if (ev.getTitle() !== e.titulo) { ev.setTitle(e.titulo); feitas++; }
        if (ev.getDescription() !== e.texto) { ev.setDescription(e.texto); feitas++; }
        if (!ev.isAllDayEvent() || !mesmoDia_(ev.getAllDayStartDate(), e.ini) || !mesmoDia_(ev.getAllDayEndDate(), fimExcl)) {
          ev.setAllDayDates(e.ini, fimExcl); feitas++;
        }
      }
      for (const uid in quero) {
        if (pendente || Date.now() - inicio > LIMITE_MS_CTO) { pendente = true; break; }
        if (vistos[uid]) continue;
        const e = quero[uid];
        const ev = cal.createAllDayEvent(e.titulo, e.ini, maisDias_(e.fim, 1), {description: e.texto});
        ev.setTag('cto', uid);
        feitas++;
        Utilities.sleep(120);
      }
      if (pendente) break;
    }
    console.log(`Agendas sincronizadas: ${feitas} alterações.${pendente ? ' Faltou terminar; a próxima rodada continua.' : ''}`);
  } finally {
    lock.releaseLock();
  }
}

// ---------- agendas ----------
function agendaCTO_(a, publicar) {
  const props = PropertiesService.getScriptProperties();
  const salvo = props.getProperty('agenda_' + a.id);
  let cal = salvo ? CalendarApp.getCalendarById(salvo) : null;
  if (!cal) {
    cal = CalendarApp.createCalendar(a.nome, {
      summary: 'Datas das olimpíadas do Centro de Treinamento Olímpico do Colégio Mackenzie Brasília. Calendário completo: ' + SITE_CTO,
      timeZone: 'America/Sao_Paulo',
    });
    props.setProperty('agenda_' + a.id, cal.getId());
    publicar = true;
  }
  if (!publicar) return cal;
  // pública só para leitura (exige o serviço "Google Calendar API" adicionado no editor)
  try {
    Calendar.Acl.insert({role: 'reader', scope: {type: 'default'}}, cal.getId());
  } catch (err) {
    if (!/already|duplicate|exist/i.test(String(err))) throw err;
  }
  registrarNaPlanilha_(a, cal.getId());
  return cal;
}

// Grava os endereços públicos na aba Listas (coluna L em diante), para o site usar.
function registrarNaPlanilha_(a, id) {
  const aba = SpreadsheetApp.getActive().getSheetByName('Listas');
  if (!aba) return;
  aba.getRange('L1:M1').setValues([['Agendas públicas (automático)', 'ID da agenda']]);
  const linha = {efai: 2, efaf: 3, em: 4}[a.id];
  aba.getRange(linha, 12, 1, 2).setValues([[a.nome, id]]);
}

// ---------- leitura da planilha (mesma lógica do site) ----------
function montarEventosCTO_() {
  const ss = SpreadsheetApp.getActive();
  const oliLinhas = linhasCTO_(ss.getSheetByName('Olimpiadas'));
  const evLinhas = linhasCTO_(ss.getSheetByName('Eventos'));
  const outrasAba = ss.getSheetByName('Treinos e outras datas');
  const outras = outrasAba ? linhasCTO_(outrasAba) : [];
  if (!oliLinhas.length || !evLinhas.length) throw new Error('Planilha vazia: nada foi alterado.');
  const oli = {};
  oliLinhas.forEach(o => { if (normCTO_(o['sigla'])) oli[keyCTO_(o['sigla'])] = o; });
  const saida = [];

  function add(sigla, rotulo, ini, fim, previsto, mes, detalhes, extra, base) {
    const o = oli[keyCTO_(sigla)] || {};
    let titulo = sigla ? `${sigla} · ${rotulo}` : rotulo;
    if (mes) titulo += ` (previsto para ${MESES_CTO[ini.getMonth()]}, dia a confirmar)`;
    else if (previsto) titulo += ' (previsto)';
    const site = normCTO_(o['site oficial']);
    const linhas = [
      normCTO_(o['nome']),
      previsto ? 'Data prevista com base no ano anterior: pode mudar. Confira sempre o site oficial.' : 'Data confirmada pelo organizador.',
      extra, detalhes,
      /^https?:\/\//i.test(site) ? 'Site oficial: ' + site : '',
      sigla ? 'Todas as datas: ' + SITE_CTO + '#' + encodeURIComponent(keyCTO_(sigla)) : 'Calendário completo: ' + SITE_CTO,
    ];
    saida.push({
      uid: `${keyCTO_(sigla)}|${base || rotulo}|${ini.getFullYear()}`,
      titulo, ini, fim: fim || ini,
      texto: linhas.filter(Boolean).join('\n'),
      segs: sigla ? segmentosCTO_(o) : ['EFAI', 'EFAF', 'EM'],
    });
  }

  evLinhas.forEach(r => {
    const sigla = normCTO_(r['sigla']);
    if (!sigla) return;
    const sit = keyCTO_(r['situacao']) || 'previsto';
    if (sit === 'cancelado') return;
    const conf = keyCTO_(r['etapas confirmadas']).split(',').map(x => x.trim()).filter(Boolean);
    const detalhes = normCTO_(r['detalhes']);
    ETAPAS_CTO.forEach(([nome, ci, cf, interno]) => {
      if (interno) return;
      let ini = dataCTO_(r[ci]); const fim = cf ? dataCTO_(r[cf]) : null;
      if (!ini && !fim) return;
      const ok = sit === 'confirmado' || conf.indexOf(keyCTO_(nome)) >= 0;
      const previsto = !ok, mes = !ok && sit.indexOf('so o mes') >= 0;
      if (nome === 'Inscrição') {
        if (ini && fim && !mesmoDia_(ini, fim)) {
          add(sigla, 'Abrem as inscrições', ini, null, previsto, mes, detalhes, 'Inscrições até ' + fmtCTO_(fim) + '.', 'insc-ini');
          add(sigla, 'Último dia de inscrição', fim, null, previsto, mes, detalhes, 'Inscrições desde ' + fmtCTO_(ini) + '.', 'insc-fim');
        } else if (fim) {
          add(sigla, 'Último dia de inscrição', fim, null, previsto, mes, detalhes, '', 'insc-fim');
        } else {
          add(sigla, 'Abrem as inscrições', ini, null, previsto, mes, detalhes, '', 'insc-ini');
        }
        return;
      }
      ini = ini || fim;
      if (fim && fim > ini && (fim - ini) / 86400000 > JANELA_CURTA_CTO) {
        add(sigla, nome + ' (início)', ini, null, previsto, mes, detalhes, 'Até ' + fmtCTO_(fim) + '.', nome + '-ini');
        add(sigla, nome + ' (último dia)', fim, null, previsto, mes, detalhes, 'Desde ' + fmtCTO_(ini) + '.', nome + '-fim');
      } else {
        add(sigla, nome, ini, fim && fim > ini ? fim : null, previsto, mes, detalhes, '', nome);
      }
    });
  });

  outras.forEach(r => {
    const ini = dataCTO_(r['data']);
    if (!ini || keyCTO_(r['visibilidade']) === 'somente equipe') return;
    const tipo = normCTO_(r['tipo']) || 'Outro';
    if (TIPOS_INTERNOS_CTO.indexOf(keyCTO_(tipo)) >= 0) return;
    const fim = dataCTO_(r['ate']);
    const extra = [normCTO_(r['horario']), normCTO_(r['local / turma'])].filter(Boolean).join(' · ');
    add(normCTO_(r['sigla']), tipo, ini, fim && fim > ini ? fim : null, false, false, normCTO_(r['detalhes']), extra,
        tipo + '-' + Utilities.formatDate(ini, 'America/Sao_Paulo', 'MMdd'));
  });
  return saida;
}

function segmentosCTO_(o) {
  let a = anoIdxCTO_(o['do ano']), b = anoIdxCTO_(o['ate o ano']);
  if (a >= 0 || b >= 0) {
    a = a >= 0 ? a : 0; b = b >= 0 ? b : 11;
    const faixas = {EFAI: [0, 4], EFAF: [5, 8], EM: [9, 11]};
    return Object.keys(faixas).filter(s => a <= faixas[s][1] && b >= faixas[s][0]);
  }
  return ['EFAI', 'EFAF', 'EM'].filter(s => keyCTO_(o[s.toLowerCase()]) === 'ok');
}
function anoIdxCTO_(v) {
  const k = keyCTO_(v), n = /(\d)/.exec(k);
  if (!n) return -1;
  return /\bem\b|serie/.test(k) ? 8 + Number(n[1]) : Number(n[1]) - 1;
}

// ---------- utilidades ----------
function linhasCTO_(aba) {
  const v = aba.getDataRange().getValues();
  if (v.length < 2) return [];
  const cab = v[0].map(keyCTO_);
  return v.slice(1).map(l => { const o = {}; cab.forEach((c, i) => o[c] = l[i]); return o; });
}
let FUSO_CTO = null;
function fusoCTO_() { return FUSO_CTO || (FUSO_CTO = SpreadsheetApp.getActive().getSpreadsheetTimeZone()); }
function normCTO_(v) { return v === null || v === undefined ? '' : String(v).trim(); }
function keyCTO_(v) { return normCTO_(v).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, ''); }
function dataCTO_(v) {
  if (v instanceof Date && !isNaN(v)) { // lê o dia no fuso da planilha, para não deslocar a data
    const p = Utilities.formatDate(v, fusoCTO_(), 'yyyy-MM-dd').split('-');
    return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
  }
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(normCTO_(v));
  return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : null;
}
function maisDias_(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function mesmoDia_(a, b) { return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate(); }
function fmtCTO_(d) { return Utilities.formatDate(d, 'America/Sao_Paulo', 'dd/MM/yyyy'); }
