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

  auxiliar.getRange(ultimaLinha,1).setValue(data); // Atribui o valor da variável "data" para a célula especificada
  auxiliar.getRange(ultimaLinha,2).setFormula('=SPLIT(A'+ultimaLinha+';"/")'); // Atribui a função "SPLIT" para a célula especificada
  auxiliar.getRange(ultimaLinha,5).setValue(tipo); // Atribui o valor da variável "tipo" para a célula especificada
  auxiliar.getRange(ultimaLinha,6).setValue(categoria); // Atribui o valor da variável "categoria" para a célula especificada
  auxiliar.getRange(ultimaLinha,7).setValue(descricao); // Atribui o valor da variável "descricao" para a célula especificada

  if (tipo == "Entrada") {
    auxiliar.getRange(ultimaLinha,8).setValue(valor); // Atribui o valor da variável "valor" para a célula especificada
  } else {
    auxiliar.getRange(ultimaLinha,8).setValue(-valor); // Atribui o valor negativo da variável "valor" para a célula especificada
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

    relatorio.getRange(i,6).setFormula("F"+(i-1)+"+E"+i+"");

  }

  SpreadsheetApp.setActiveSheet(relatorio);
  SpreadsheetApp.getUi().alert("Relatório gerado com sucesso!", 'Após visualizá-lo, retorne para aba "Gerador de relatórios" para que possa enviá-lo por e-mail para seus destinatários.', SpreadsheetApp.getUi().ButtonSet.OK);

}


// ENVIA O RELATÓRIO POR E-MAIL (NO FORMATO PDF)

var destinatario = gerador.getRange("K4:K5").getValue();
var mensagem = gerador.getRange("I4:I6").getValue();

var email = {
  to: destinatario,
  subject: "Relatório Financeiro",
  body: mensagem,
  name: "Thullyo Damasceno",
  attachments: [planilha.getAs(MimeType.PDF).setName("Relatório Financeiro"+".pdf")]
}

function enviar() {

  if(Browser.msgBox('Compartilhar "Relatório"','Deseja compartilhar o relatório financeiro com "'+destinatario+'"?', Browser.Buttons.YES_NO) == 'yes') {

    cadastro.hideSheet();
    movimentacoes.hideSheet();
    gerador.hideSheet();

    MailApp.sendEmail(email);

    cadastro.showSheet();
    movimentacoes.showSheet();
    gerador.showSheet();

  }

}
