var planilha = SpreadsheetApp.getActiveSpreadsheet(); // Acessa a planilha
var cadastro = planilha.getSheetByName("Cadastro"); // Acessa a aba "Cadastro"
var auxiliar = planilha.getSheetByName("Auxiliar"); // Acessa a aba "Auxiliar"
var movimentacoes = planilha.getSheetByName("Movimentações"); // Acessa a aba "Movimentações"
var gerador = planilha.getSheetByName("Gerador de relatórios"); // Acessa a aba "Gerador de relatórios"
var relatorio = planilha.getSheetByName("Relatório"); // Acessa a aba "Relatório"

var data = cadastro.getRange("C3:G3").getValue(); // Armazena o valor do campo "data"
var tipo = cadastro.getRange("C5").getValue(); // Armazena o valor do campo "tipo"
var categoria = cadastro.getRange("F5:G5").getValue(); // Armazena o valor do campo "categoria"
var descricao = cadastro.getRange("C7:G7").getValue(); // Armazena o valor do campo "descrição"
var valor = cadastro.getRange("C9:G9").getValue(); // Armazena o valor do campo "valor"


// CADASTRA AS MOVIMENTAÇÕES FINANCEIRAS 

function cadastrar() {

  var ultimaLinha = auxiliar.getLastRow()+1; // Seleciona a linha que fica logo após a última linha da aba "Auxiliar"

  auxiliar.getRange(ultimaLinha,1).setValue(data);
  auxiliar.getRange(ultimaLinha,2).setFormula('=SPLIT(A'+ultimaLinha+';"/")');
  auxiliar.getRange(ultimaLinha,5).setValue(tipo);
  auxiliar.getRange(ultimaLinha,6).setValue(categoria);
  auxiliar.getRange(ultimaLinha,7).setValue(descricao);

  if (tipo == "Entrada") {
    auxiliar.getRange(ultimaLinha,8).setValue(valor);
  } else {
    auxiliar.getRange(ultimaLinha,8).setValue(-valor);
  }

  if (ultimaLinha != 2) {
    movimentacoes.getRange(ultimaLinha,9).setFormula("I"+(ultimaLinha-1)+"+H"+ultimaLinha+"");
  } else {
    movimentacoes.getRange(ultimaLinha,9).setFormula("H2");
  }
  
  limparCampos();

}


// LIMPA OS CAMPOS DA ABA "Cadastro" 

function limparCampos() {

  cadastro.getRange("C3:G3").clearContent(); // Limpa o conteúdo do intervalo "C3:G3"
  cadastro.getRange("C5").clearContent(); // Limpa o conteúdo da célula "C5"
  cadastro.getRange("F5:G5").clearContent(); // Limpa o conteúdo do intervalo "F5:G5"
  cadastro.getRange("C7:G7").clearContent(); // Limpa o conteúdo do intervalo "C7:G7"
  cadastro.getRange("C9:G9").clearContent(); // Limpa o conteúdo do intervalo "C9:G9"

}


// GERA O RELATÓRIO

function gerar() {

  relatorio.getRange("F2:F").clearContent();
  relatorio.getRange("F2").setFormula("E2");

  for (var i = 3; i <= relatorio.getLastRow(); i++) {

    relatorio.getRange(i,6).setFormula('IF(F'+(i-1)+'+E'+i+'<>0; F'+(i-1)+'+E'+i+'; "")');

  }

  SpreadsheetApp.getUi().alert("Relatório gerado com sucesso!", 'Após visualizá-lo, retorne para a aba "Gerador de relatórios" para que possa enviá-lo por e-mail para seus destinatários.', SpreadsheetApp.getUi().ButtonSet.OK);
  SpreadsheetApp.setActiveSheet(relatorio);

}


// ENVIA O RELATÓRIO POR E-MAIL (NO FORMATO PDF)

function enviar() {

  var destinatario = gerador.getRange("J4:J5").getValue();
  var mensagem = gerador.getRange("H4:H6").getValue();

  var email = {
    to: destinatario,
    subject: "Relatório Financeiro",
    body: mensagem,
    name: "Thullyo Damasceno",
    attachments: [planilha.getAs(MimeType.PDF).setName("Relatório Financeiro"+".pdf")]
  }

  if(Browser.msgBox("Deseja compartilhar o relatório financeiro com "+destinatario+"?", Browser.Buttons.YES_NO) == 'yes') {

    cadastro.hideSheet();
    movimentacoes.hideSheet();
    gerador.hideSheet();
    relatorio.dele

    MailApp.sendEmail(email);

    cadastro.showSheet();
    movimentacoes.showSheet();
    gerador.showSheet();

  }

}
