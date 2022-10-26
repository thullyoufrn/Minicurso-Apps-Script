var planilha = SpreadsheetApp.getActiveSpreadsheet(); // Acessa a planilha
var cadastro = planilha.getSheetByName("Cadastro"); // Acessa a aba "Cadastro"
var movOcultas = planilha.getSheetByName("Movimentações ocultas"); // Acessa a aba "Movimentações ocultas"
var relatorio = planilha.getSheetByName("Relatório"); // Acessa a aba "Relatório"

var data = cadastro.getRange("C3:G3").getValue(); // Armazena o valor do campo "data"
var tipo = cadastro.getRange("C5").getValue(); // Armazena o valor do campo "tipo"
var categoria = cadastro.getRange("F5:G5").getValue(); // Armazena o valor do campo "categoria"
var descricao = cadastro.getRange("C7:G7").getValue(); // Armazena o valor do campo "descrição"
var valor = cadastro.getRange("C9:G9").getValue(); // Armazena o valor do campo "valor"

////////////////////////////////////////////
// CADASTRAR AS MOVIMENTAÇÕES FINANCEIRAS //
////////////////////////////////////////////

function cadastrar() {

  var ultimaLinha = movOcultas.getLastRow()+1; // Adiciona uma linha na aba "Movimentações ocultas" e seleciona a linha adicionada

  var dataSeparada = categoria.split(" ");

  Logger.log(""+data+"");

  /* movOcultas.getRange(ultimaLinha,1).setValue(dataSeparada[0]);
  movOcultas.getRange(ultimaLinha,1).setValue(dataSeparada[1]);
  movOcultas.getRange(ultimaLinha,1).setValue(dataSeparada[2]);
  movOcultas.getRange(ultimaLinha,4).setValue(tipo);
  movOcultas.getRange(ultimaLinha,5).setValue(categoria);
  movOcultas.getRange(ultimaLinha,6).setValue(descricao);
  movOcultas.getRange(ultimaLinha,7).setValue(valor); */
  
  // limparCampos();

}

////////////////////////////////////////
// LIMPAR OS CAMPOS DA ABA "Cadastro" //
////////////////////////////////////////

function limparCampos() {

  cadastro.getRange("C3:G3").clearContent(); // Limpa o conteúdo do intervalo "C3:G3"
  cadastro.getRange("C5").clearContent(); // Limpa o conteúdo da célula "C5"
  cadastro.getRange("F5:G5").clearContent(); // Limpa o conteúdo do intervalo "F5:G5"
  cadastro.getRange("C7:G7").clearContent(); // Limpa o conteúdo do intervalo "C7:G7"
  cadastro.getRange("C9:G9").clearContent(); // Limpa o conteúdo do intervalo "C9:G9"

}

////////////////////////////
// VISUALIZAR O RELATÓRIO //
////////////////////////////

function ver() {

  SpreadsheetApp.setActiveSheet(relatorio);

}

//////////////////////////////////////////////////
// ENVIAR O RELATÓRIO NO FORMATO PDF VIA E-MAIL //
//////////////////////////////////////////////////

function enviar() {

  SpreadsheetApp.getUi().prompt(
    "Você está prestes a compartilhar este relatório",
    "Digite o e-mail do destinatário:", 
    SpreadsheetApp.getUi().ButtonSet.OK_CANCEL);

}
