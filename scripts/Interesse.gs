// Calendário Olímpico CTO – formulário "Tenho interesse".
//
// Cria um Google Forms para alunos e famílias demonstrarem interesse numa olimpíada.
// As respostas vão para uma planilha PRIVADA na conta do CTO (não compartilhada).
// Quem preenche recebe um e-mail de confirmação.
// A lista de olimpíadas do formulário acompanha a aba Olimpiadas sozinha (todo dia).
//
// Instalação (uma vez, logado só na conta do CTO):
//   1. No mesmo projeto do Apps Script da planilha: + ao lado de "Arquivos" → Script → nome "Interesse" → cole este código.
//   2. Escolha a função instalarFormulario no menu de cima e clique em Executar. Autorize.

const FORM_TITULO_CTO = 'CTO Mackenzie Brasília · Tenho interesse em uma olimpíada';
const FORM_SEP_CTO = ' – '; // separa sigla e nome na lista; o site usa o mesmo separador
const SERIES_CTO = ['1º ano EF','2º ano EF','3º ano EF','4º ano EF','5º ano EF','6º ano EF','7º ano EF','8º ano EF','9º ano EF','1ª série EM','2ª série EM','3ª série EM'];
// true = quem preencher recebe um e-mail de confirmação (enviado pela conta do CTO)
const CONFIRMAR_PARA_QUEM_PREENCHEU = true;
// true = a conta do CTO também recebe um e-mail a cada novo interesse
const AVISAR_CTO = false;

// Rode UMA vez: cria o formulário, a planilha privada de respostas, os avisos e a atualização diária.
function instalarFormulario() {
  const props = PropertiesService.getScriptProperties();
  let form = props.getProperty('form_id') ? FormApp.openById(props.getProperty('form_id')) : null;
  if (!form) {
    form = FormApp.create(FORM_TITULO_CTO);
    props.setProperty('form_id', form.getId());
    montarPerguntas_(form);
    const respostas = SpreadsheetApp.create('CTO · Interessados nas olimpíadas (PRIVADO)');
    respostas.setSpreadsheetLocale('pt_BR');
    respostas.setSpreadsheetTimeZone('America/Sao_Paulo');
    form.setDestination(FormApp.DestinationType.SPREADSHEET, respostas.getId());
    props.setProperty('respostas_id', respostas.getId());
    SpreadsheetApp.flush();
    organizarPlanilhaInteressados();
  }
  atualizarListaOlimpiadas();
  ScriptApp.getProjectTriggers()
    .filter(t => ['avisarNovoInteresse', 'atualizarListaOlimpiadas'].indexOf(t.getHandlerFunction()) >= 0)
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('avisarNovoInteresse').forForm(form).onFormSubmit().create();
  ScriptApp.newTrigger('atualizarListaOlimpiadas').timeBased().everyDays(1).atHour(5).create();
  registrarFormNaPlanilha_(form);
  console.log('Formulário: ' + form.getPublishedUrl());
  console.log('Respostas (privado): ' + SpreadsheetApp.openById(props.getProperty('respostas_id')).getUrl());
}

function montarPerguntas_(form) {
  form.setDescription(
    'Quer participar de uma olimpíada com o Centro de Treinamento Olímpico? Preencha e a equipe do CTO entra em contato.\n' +
    'Este formulário registra só o seu interesse: a inscrição oficial vem depois, com a equipe.\n' +
    'Os dados são usados apenas pela equipe do CTO para falar sobre a olimpíada escolhida e não são compartilhados.')
    .setCollectEmail(false)
    .setLimitOneResponsePerUser(false)
    .setAllowResponseEdits(false)
    .setShowLinkToRespondAgain(true)
    .setProgressBar(false)
    .setConfirmationMessage('Recebemos o seu interesse! A equipe do CTO vai entrar em contato pelo e-mail informado. ' +
      'Quer participar de outra olimpíada? Use o link "Enviar outra resposta".');

  form.addListItem().setTitle('Olimpíada').setHelpText('Se veio pelo site, ela já está marcada.').setRequired(true)
    .setChoiceValues(['(carregando)']);
  form.addTextItem().setTitle('Nome completo do aluno').setRequired(true);
  form.addMultipleChoiceItem().setTitle('Série').setChoiceValues(SERIES_CTO).setRequired(true);
  form.addTextItem().setTitle('Turma').setHelpText('Ex.: A, B, 7ºC').setRequired(false);
  form.addMultipleChoiceItem().setTitle('Quem está preenchendo?').setChoiceValues(['O próprio aluno', 'Responsável']).setRequired(true);
  form.addTextItem().setTitle('E-mail para contato').setRequired(true)
    .setValidation(FormApp.createTextValidation().requireTextIsEmail().setHelpText('Digite um e-mail válido.').build());
  form.addTextItem().setTitle('WhatsApp (opcional)').setHelpText('Com DDD. Ex.: (61) 99999-9999').setRequired(false);
  form.addParagraphTextItem().setTitle('Alguma observação? (opcional)').setRequired(false);
  form.addCheckboxItem().setTitle('Autorização').setRequired(true)
    .setChoiceValues(['Autorizo a equipe do CTO a usar estes dados apenas para contato sobre esta olimpíada.']);
}

// Troca a pergunta "Série" (menu, que trava no celular) por opções para tocar e refaz a planilha
// de respostas na ordem certa. Só funciona enquanto houver no máximo 3 respostas (as de teste),
// que são apagadas. Depois organiza a planilha para a equipe.
function corrigirSerie() {
  const props = PropertiesService.getScriptProperties();
  const form = FormApp.openById(props.getProperty('form_id'));
  if (form.getResponses().length > 3) throw new Error('Já há respostas reais: nada foi alterado. Fale com quem montou o sistema.');
  const antigo = form.getItems(FormApp.ItemType.LIST).filter(i => i.getTitle() === 'Série')[0];
  if (antigo) {
    const pos = antigo.getIndex();
    const novo = form.addMultipleChoiceItem().setTitle('Série').setChoiceValues(SERIES_CTO).setRequired(true);
    form.moveItem(novo.getIndex(), pos);
    form.deleteItem(pos + 1);
  }
  const ss = SpreadsheetApp.openById(props.getProperty('respostas_id'));
  form.removeDestination();
  form.deleteAllResponses();
  SpreadsheetApp.flush();
  const temp = ss.getSheetByName('_temp') || ss.insertSheet('_temp');
  ss.getSheets().filter(s => s.getName() !== '_temp').forEach(s => ss.deleteSheet(s));
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  Utilities.sleep(3000);
  ss.deleteSheet(temp);
  organizarPlanilhaInteressados();
}

// Mantém a lista de olimpíadas do formulário igual à aba Olimpiadas (roda todo dia às 5h).
function atualizarListaOlimpiadas() {
  const form = FormApp.openById(PropertiesService.getScriptProperties().getProperty('form_id'));
  const aba = SpreadsheetApp.getActive().getSheetByName('Olimpiadas');
  const v = aba.getDataRange().getValues();
  const cab = v[0].map(x => String(x).trim().toLowerCase());
  const iS = cab.indexOf('sigla'), iN = cab.indexOf('nome');
  const opcoes = v.slice(1)
    .filter(l => String(l[iS]).trim())
    .map(l => opcaoOlimpiada_(l[iS], l[iN]))
    .sort((a, b) => a.localeCompare(b, 'pt-BR'));
  const item = perguntaOlimpiada_(form);
  if (opcoes.length) item.setChoiceValues(opcoes);
}

function opcaoOlimpiada_(sigla, nome) {
  sigla = String(sigla).trim(); nome = String(nome || '').trim();
  return nome ? sigla + FORM_SEP_CTO + nome : sigla;
}
function perguntaOlimpiada_(form) {
  return form.getItems(FormApp.ItemType.LIST).filter(i => i.getTitle() === 'Olimpíada')[0].asListItem();
}

// A cada resposta: confirmação para quem preencheu e, se ligado, aviso para a conta do CTO.
function avisarNovoInteresse(e) {
  const resp = {};
  e.response.getItemResponses().forEach(r => {
    const v = r.getResponse();
    resp[r.getItem().getTitle()] = Array.isArray(v) ? v.join(', ') : v;
  });
  const props = PropertiesService.getScriptProperties();
  const planilha = 'https://docs.google.com/spreadsheets/d/' + props.getProperty('respostas_id') + '/edit';
  const linhas = ['Olimpíada', 'Nome completo do aluno', 'Série', 'Turma', 'Quem está preenchendo?', 'E-mail para contato', 'WhatsApp (opcional)', 'Alguma observação? (opcional)']
    .filter(k => resp[k]).map(k => k.replace(' (opcional)', '') + ': ' + resp[k]);
  if (AVISAR_CTO) MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: 'Novo interesse: ' + (resp['Olimpíada'] || 'olimpíada') + ' · ' + (resp['Nome completo do aluno'] || ''),
    body: linhas.join('\n') + '\n\nLista completa (privada): ' + planilha,
    name: 'Calendário Olímpico CTO',
  });
  const contato = resp['E-mail para contato'];
  if (CONFIRMAR_PARA_QUEM_PREENCHEU && contato) {
    MailApp.sendEmail({
      to: contato,
      subject: 'Recebemos seu interesse: ' + (resp['Olimpíada'] || 'olimpíada'),
      body: 'Olá!\n\nRecebemos o interesse de ' + (resp['Nome completo do aluno'] || 'o aluno') + ' em participar desta olimpíada:\n' +
        (resp['Olimpíada'] || '') + '\n\nA equipe do Centro de Treinamento Olímpico vai entrar em contato com as próximas orientações. ' +
        'Este e-mail confirma só o interesse: a inscrição oficial é feita depois, com a equipe.\n\n' +
        'Calendário das olimpíadas: https://ctomackenziebrasilia.github.io/\n\nCTO · Colégio Mackenzie Brasília',
      name: 'CTO Mackenzie Brasília',
    });
  }
}

// Organiza a planilha PRIVADA de interessados para a equipe (pode rodar de novo quando quiser:
// não apaga respostas nem anotações). Cria as abas "Resumo", "Interessados" e "Como usar".
function organizarPlanilhaInteressados() {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('respostas_id'));
  const BORDO = '#A8192E', TINTA = '#1F1F1F', CINZA = '#6B6B70', CLARO = '#F6F3F1';
  ss.setSpreadsheetLocale('pt_BR');

  // --- Interessados (respostas do formulário + 2 colunas da equipe) ---
  const resp = ss.getSheets().filter(s => s.getFormUrl())[0];
  if (!resp) throw new Error('Não achei a aba de respostas do formulário.');
  resp.setName('Interessados');
  const nForm = resp.getLastColumn() >= 10 ? 10 : resp.getLastColumn();
  const cab = resp.getRange(1, 1, 1, resp.getMaxColumns()).getValues()[0].map(String);
  let colStatus = cab.indexOf('Contato (equipe)') + 1, colNotas = cab.indexOf('Anotações da equipe') + 1;
  if (!colStatus) { colStatus = Math.max(resp.getLastColumn(), nForm) + 1; ensureCols_(resp, colStatus); resp.getRange(1, colStatus).setValue('Contato (equipe)'); }
  if (!colNotas) { colNotas = colStatus + 1; ensureCols_(resp, colNotas); resp.getRange(1, colNotas).setValue('Anotações da equipe'); }
  const ultCol = Math.max(colNotas, resp.getLastColumn());
  resp.setFrozenRows(1); resp.setFrozenColumns(3);
  resp.getRange(1, 1, 1, ultCol).setBackground(BORDO).setFontColor('#FFFFFF').setFontWeight('bold')
    .setFontFamily('Roboto').setVerticalAlignment('middle').setWrap(true);
  resp.getRange(1, colStatus, 1, 2).setBackground('#5C2A32');
  resp.setRowHeight(1, 44);
  resp.getRange(2, 1, resp.getMaxRows() - 1, ultCol).setFontFamily('Roboto').setVerticalAlignment('middle');
  resp.getRange(2, 1, resp.getMaxRows() - 1, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  const larg = {1: 125, 2: 280, 3: 220, 4: 105, 5: 70, 6: 135, 7: 220, 8: 140, 9: 220, 10: 90};
  Object.keys(larg).forEach(c => { if (Number(c) <= nForm) resp.setColumnWidth(Number(c), larg[c]); });
  resp.setColumnWidth(colStatus, 150); resp.setColumnWidth(colNotas, 260);
  if (nForm >= 10) resp.hideColumns(10); // "Autorização": sempre marcada (obrigatória)
  resp.getRange(2, colStatus, resp.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['A contatar', 'Contatado', 'Inscrito', 'Não vai participar'], true)
      .setAllowInvalid(false).setHelpText('Atualize depois de falar com a família.').build());
  // cores por situação do contato (vazio = novo, ainda não contatado)
  const linhaInteira = resp.getRange(2, 1, resp.getMaxRows() - 1, ultCol);
  const L = colLetra_(colStatus);
  const regras = [
    ['=AND($B2<>"";$' + L + '2="")', '#FFF4D6', '#7A5600'],
    ['=$' + L + '2="A contatar"', '#FFF4D6', '#7A5600'],
    ['=$' + L + '2="Contatado"', '#E7F1FB', '#1F4E86'],
    ['=$' + L + '2="Inscrito"', '#E5F4E4', '#1E6B3A'],
    ['=$' + L + '2="Não vai participar"', '#F0EFEC', '#8A8A8A'],
  ].map(([f, bg, fg]) => SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(f)
    .setBackground(bg).setFontColor(fg).setRanges([resp.getRange(2, colStatus, resp.getMaxRows() - 1, 1)]).build());
  regras.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=$' + L + '2="Não vai participar"')
    .setFontColor('#9A9A9A').setRanges([linhaInteira]).build());
  resp.setConditionalFormats(regras);
  resp.getBandings().forEach(b => b.remove());
  if (resp.getFilter()) resp.getFilter().remove();
  resp.getRange(1, 1, resp.getMaxRows(), ultCol).createFilter();
  resp.setTabColor(BORDO);
  resp.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(p => p.remove());
  resp.getRange(1, 1, resp.getMaxRows(), nForm).protect()
    .setDescription('Colunas preenchidas pelo formulário: evite editar').setWarningOnly(true);

  // --- Resumo ---
  let res = ss.getSheetByName('Resumo') || ss.getSheetByName('Por olimpíada') || ss.insertSheet('Resumo');
  res.setName('Resumo'); res.clear(); res.clearConditionalFormatRules();
  res.setHiddenGridlines(true); res.setTabColor(TINTA);
  res.getRange('A1:F60').setFontFamily('Roboto').setFontColor(TINTA).setVerticalAlignment('middle');
  res.setColumnWidth(1, 20);
  [300, 140, 140, 140, 150].forEach((w, i) => res.setColumnWidth(i + 2, w));
  res.getRange('B2').setValue('Interessados nas olimpíadas').setFontSize(20).setFontWeight('bold');
  res.getRange('B3').setValue('Dados de alunos: uso exclusivo da equipe do CTO. Não compartilhe esta planilha.')
    .setFontColor(BORDO).setFontSize(10).setFontWeight('bold');
  const I = "Interessados!", S = I + L + "2:" + L;
  const cards = [
    ['Interessados', '=COUNTA(' + I + 'B2:B)'],
    ['Novos (7 dias)', '=COUNTIF(' + I + 'A2:A;">="&TODAY()-7)'],
    ['A contatar', '=COUNTIFS(' + I + 'B2:B;"<>";' + S + ';"")+COUNTIF(' + S + ';"A contatar")'],
    ['Inscritos', '=COUNTIF(' + S + ';"Inscrito")'],
  ];
  cards.forEach(([rot, f], i) => {
    const c = 2 + i + (i > 0 ? 0 : 0);
    res.getRange(5, c).setValue(rot).setFontSize(9).setFontColor(CINZA).setFontWeight('bold');
    res.getRange(6, c).setFormula(f).setFontSize(24).setFontWeight('bold').setHorizontalAlignment('left');
    res.getRange(5, c, 2, 1).setBackground(CLARO);
  });
  res.getRange(6, 4).setFontColor('#7A5600');
  res.setRowHeight(6, 46);
  res.getRange('B8').setValue('Por olimpíada').setFontSize(13).setFontWeight('bold');
  res.getRange('B9').setFormula('=IFERROR(QUERY(' + I + 'A2:B;"select B, count(A), max(A) where B is not null group by B order by count(A) desc label B \'Olimpíada\', count(A) \'Interessados\', max(A) \'Último interesse\' format max(A) \'dd/mm/yyyy\'";0);"Ainda não há respostas.")');
  res.getRange('E9').setValue('A contatar');
  res.getRange('F9').setValue('Inscritos');
  res.getRange('E10').setFormula('=BYROW(B10:B;LAMBDA(o;IF(o="";"";COUNTIFS(' + I + 'B2:B;o;' + S + ';"")+COUNTIFS(' + I + 'B2:B;o;' + S + ';"A contatar"))))');
  res.getRange('F10').setFormula('=BYROW(B10:B;LAMBDA(o;IF(o="";"";COUNTIFS(' + I + 'B2:B;o;' + S + ';"Inscrito"))))');
  res.getRange('B9:F9').setBackground(BORDO).setFontColor('#FFFFFF').setFontWeight('bold');
  res.getRange('C9:F60').setHorizontalAlignment('center');
  res.setFrozenRows(0);
  res.setConditionalFormats([
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($B10<>"";ISEVEN(ROW()))')
      .setBackground(CLARO).setRanges([res.getRange('B10:F60')]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
      .setFontColor('#7A5600').setBold(true).setRanges([res.getRange('E10:E60')]).build(),
  ]);

  // --- Como usar ---
  let ajuda = ss.getSheetByName('Como usar') || ss.insertSheet('Como usar');
  ajuda.clear(); ajuda.setHiddenGridlines(true); ajuda.setTabColor(CINZA);
  ajuda.setColumnWidth(1, 20); ajuda.setColumnWidth(2, 760);
  const textos = [
    ['Como usar esta planilha', 18, true],
    ['', 10, false],
    ['1.  Cada linha da aba "Interessados" é um pedido feito pelo formulário do site (botão "Tenho interesse" na ficha da olimpíada).', 11, false],
    ['2.  Linhas em amarelo ainda não foram contatadas. Depois de falar com a família, escolha a situação na coluna "Contato (equipe)": Contatado, Inscrito ou Não vai participar.', 11, false],
    ['3.  Use a coluna "Anotações da equipe" para registrar combinados (ex.: "pais confirmaram por WhatsApp em 10/03").', 11, false],
    ['4.  Para ver só uma olimpíada: na aba "Interessados", clique no ícone de filtro do título "Olimpíada" e marque a desejada.', 11, false],
    ['5.  A aba "Resumo" mostra os totais e quantos ainda faltam contatar em cada olimpíada. Ela se atualiza sozinha.', 11, false],
    ['6.  Não edite as colunas que vêm do formulário (até "Alguma observação?"). Não apague a aba "Interessados" nem mude a ordem das colunas.', 11, false],
    ['', 10, false],
    ['Privacidade: são dados de alunos, em grande parte menores de idade. Use só para o contato sobre a olimpíada escolhida e não compartilhe esta planilha fora da equipe do CTO. Para apagar os dados de alguém que pedir, apague a linha inteira.', 10, true],
  ];
  textos.forEach(([txt, tam, neg], i) => ajuda.getRange(2 + i, 2).setValue(txt).setFontSize(tam).setFontWeight(neg ? 'bold' : 'normal')
    .setFontFamily('Roboto').setWrap(true).setFontColor(i === textos.length - 1 ? BORDO : TINTA));

  // ordem das abas e limpeza
  ss.setActiveSheet(res); ss.moveActiveSheet(1);
  ss.setActiveSheet(resp); ss.moveActiveSheet(2);
  ss.setActiveSheet(ajuda); ss.moveActiveSheet(3);
  ['Página1', 'Sheet1'].forEach(n => { const s = ss.getSheetByName(n); if (s && ss.getSheets().length > 3) ss.deleteSheet(s); });
  ss.setActiveSheet(res);
  console.log('Planilha de interessados organizada: ' + ss.getUrl());
}
function ensureCols_(sh, n) { if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns()); }
function colLetra_(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }

// Grava na aba Listas (colunas O:P) o endereço público do formulário e o campo usado para
// já marcar a olimpíada. Não há dados de alunos aqui: é só o link que o site usa no botão.
function registrarFormNaPlanilha_(form) {
  const item = perguntaOlimpiada_(form);
  const exemplo = item.getChoices()[0].getValue();
  const url = form.createResponse().withItemResponse(item.createResponse(exemplo)).toPrefilledUrl();
  const campo = (/[?&](entry\.\d+)=/.exec(url) || [])[1] || '';
  const aba = SpreadsheetApp.getActive().getSheetByName('Listas');
  aba.getRange('O1:P3').setValues([
    ['Formulário "Tenho interesse" (automático)', ''],
    ['Endereço', form.getPublishedUrl()],
    ['Campo da olimpíada', campo],
  ]);
}
