function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    // AÇÃO 1: CRIAR NOVO PEDIDO
    if (data.action === "createOrder") {
      var logoUrl = "Sem logo";

      if (data.logoBase64 && data.logoBase64 !== "") {
        var blob = Utilities.newBlob(Utilities.base64Decode(data.logoBase64), data.mimeType, data.fileName);
        var folderName = "Logos_Plataforma_Sites";
        var folders = DriveApp.getFoldersByName(folderName);
        var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
        folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        var file = folder.createFile(blob);
        logoUrl = file.getUrl();
      }

      // Constrói o Prompt incluindo os Presets selecionados
      var promptStr = 'Atue como um programador web especialista. Crie uma landing page responsiva num único ficheiro index.html com CSS e JS integrados. Cores: Principal ' + data.primaryColor + ', Secundária ' + data.secondaryColor + '. Fonte: ' + data.font + '. Logo: ' + logoUrl + '. Recursos obrigatórios: [' + data.presets + ']. Descrição detalhada: "' + data.description + '". Retorne apenas o código completo e estruturado.';

      sheet.appendRow([
        new Date(), data.siteName, data.primaryColor, data.secondaryColor, data.font, data.presets, data.description, logoUrl, "Pendente", promptStr, "[]"
      ]);

      return ContentService.createTextOutput(JSON.stringify({"status": "sucesso"})).setMimeType(ContentService.MimeType.JSON);
    }

    // AÇÃO 2: ADICIONAR MENSAGEM / AJUSTE DE CHAT
    if (data.action === "addAdjustment") {
      var rows = sheet.getDataRange().getValues();
      for (var i = 1; i < rows.length; i++) {
        if (rows[i][1].toString().toLowerCase() === data.siteName.toLowerCase()) {
          var currentHistory = [];
          try { currentHistory = JSON.parse(rows[i][10] || "[]"); } catch(err) {}
          
          currentHistory.push({ autor: "user", texto: data.adjustmentText, data: new Date() });
          
          // Atualiza o histórico na coluna K (índice 11) e altera status para Ajuste Solicitado
          sheet.getRange(i + 1, 11).setValue(JSON.stringify(currentHistory));
          sheet.getRange(i + 1, 9).setValue("Ajuste Solicitado");
          
          return ContentService.createTextOutput(JSON.stringify({"status": "sucesso"})).setMimeType(ContentService.MimeType.JSON);
        }
      }
    }

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "erro", "mensagem": error.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// CONSULTA DE STATUS E CHAT VIA GET
function doGet(e) {
  var siteQuery = e.parameter.siteName;
  if (!siteQuery) return ContentService.createTextOutput(JSON.stringify({"status": "erro"})).setMimeType(ContentService.MimeType.JSON);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rows = sheet.getDataRange().getValues();

  for (var i = 1; i < rows.length; i++) {
    if (rows[i][1].toString().toLowerCase() === siteQuery.toLowerCase()) {
      var history = [];
      try { history = JSON.parse(rows[i][10] || "[]"); } catch(err) {}
      
      return ContentService.createTextOutput(JSON.stringify({
        "status": "encontrado",
        "siteName": rows[i][1],
        "statusPedido": rows[i][8],
        "historico": history
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput(JSON.stringify({"status": "nao_encontrado"})).setMimeType(ContentService.MimeType.JSON);
}
