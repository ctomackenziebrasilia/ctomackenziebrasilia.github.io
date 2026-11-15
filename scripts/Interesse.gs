// Calendário Olímpico CTO – formulário "Tenho interesse".
//
// Cria um Google Forms para alunos e famílias demonstrarem interesse numa olimpíada.
// As respostas vão para uma planilha PRIVADA na conta do CTO (não compartilhada) e
// a conta do CTO recebe um e-mail a cada novo interesse.
// A lista de olimpíadas do formulário acompanha a aba Olimpiadas sozinha (todo dia).
//
// Instalação (uma vez, logado só na conta do CTO):
//   1. No mesmo projeto do Apps Script da planilha: + ao lado de "Arquivos" → Script → nome "Interesse" → cole este código.
//   2. Escolha a função instalarFormulario no menu de cima e clique em Executar. Autorize.

const FORM_TITULO_CTO = 'CTO Mackenzie Brasília · Tenho interesse em uma olimpíada';
const FORM_SEP_CTO = ' – '; // separa sigla e nome na lista; o site usa o mesmo separador
const SERIES_CTO = ['1º ano EF','2º ano EF','3º ano EF','4º ano EF','5º ano EF','6º ano EF','7º ano EF','8º ano EF','9º ano EF','1ª série EM','2ª série EM','3ª série EM'];
// true = quem preencher também recebe um e-mail de confirmação (sai da conta do CTO)
const CONFIRMAR_PARA_QUEM_PREENCHEU = false;

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
    montarResumo_(respostas);
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
  form.addListItem().setTitle('Série').setChoiceValues(SERIES_CTO).setRequired(true);
  form.addTextItem().setTitle('Turma').setHelpText('Ex.: A, B, 7ºC').setRequired(false);
  form.addMultipleChoiceItem().setTitle('Quem está preenchendo?').setChoiceValues(['O próprio aluno', 'Responsável']).setRequired(true);
  form.addTextItem().setTitle('E-mail para contato').setRequired(true)
    .setValidation(FormApp.createTextValidation().requireTextIsEmail().setHelpText('Digite um e-mail válido.').build());
  form.addTextItem().setTitle('WhatsApp (opcional)').setHelpText('Com DDD. Ex.: (61) 99999-9999').setRequired(false);
  form.addParagraphTextItem().setTitle('Alguma observação? (opcional)').setRequired(false);
  form.addCheckboxItem().setTitle('Autorização').setRequired(true)
    .setChoiceValues(['Autorizo a equipe do CTO a usar estes dados apenas para contato sobre esta olimpíada.']);
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

// Aviso por e-mail para a conta do CTO a cada novo interesse.
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
  MailApp.sendEmail({
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
      body: 'Olá!\n\nRecebemos o interesse de ' + (resp['Nome completo do aluno'] || 'o aluno') + ' em participar da ' +
        (resp['Olimpíada'] || 'olimpíada') + '. A equipe do Centro de Treinamento Olímpico vai entrar em contato com as próximas orientações.\n\n' +
        'Calendário das olimpíadas: https://ctomackenziebrasilia.github.io/\n\nCTO · Colégio Mackenzie Brasília',
      name: 'CTO Mackenzie Brasília',
    });
  }
}

// Aba "Por olimpíada" na planilha privada: quantos interessados em cada uma.
function montarResumo_(ss) {
  const resp = ss.getSheets().filter(s => s.getFormUrl())[0] || ss.getSheets()[0];
  let resumo = ss.getSheetByName('Por olimpíada') || ss.insertSheet('Por olimpíada');
  resumo.getRange('A1').setFormula(
    "=IFERROR(QUERY('" + resp.getName() + "'!A:J;\"select B, count(A) where B is not null group by B order by count(A) desc label B 'Olimpíada', count(A) 'Interessados'\";1);\"Ainda não há respostas.\")");
  resumo.setColumnWidth(1, 420);
  const sobra = ss.getSheetByName('Página1') || ss.getSheetByName('Sheet1');
  if (sobra && ss.getSheets().length > 2) ss.deleteSheet(sobra);
}

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
