
function onOpen() {
    SpreadsheetApp.getUi()
      .createMenu('Import On-Premises Data')
      .addItem('Configure OData', 'showDialog')
      .addItem('Clear Content', 'clearContent')
      .addToUi();
}

//Ask For OData Config
function showDialog() {
  var html = HtmlService.createHtmlOutputFromFile('ODataConfigDialog')
      .setWidth(500)
      .setHeight(500);
  SpreadsheetApp.getUi()
      .showModalDialog(html, 'Configure OData');
}

function clearContent(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  sheet.clearContents();
}

//Process the OData Configuration
function itemAdd(form)
{
  var odataurl = form.odataurl;

  odataurl = odataurl.trim();

  var lastChar = odataurl.charAt(odataurl.length - 1);
  if (lastChar != '/') {
    odataurl = odataurl + '/';
  }
  
  PropertiesService.getScriptProperties().setProperty('odataurl', odataurl);
  PropertiesService.getScriptProperties().setProperty('username', form.username);
  PropertiesService.getScriptProperties().setProperty('password', form.password);

  getMetadata();
}

//Fetches Metadata
function getMetadata()
{
  var options = {};
  var metadata = {};
  var odataurl = PropertiesService.getScriptProperties().getProperty('odataurl');
  var username = PropertiesService.getScriptProperties().getProperty('username');
  var password = PropertiesService.getScriptProperties().getProperty('password');
  
  var odataMetadataLink = odataurl + '$metadata';
  options.headers = {"Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)};
  var responseMetadata = UrlFetchApp.fetch(odataMetadataLink, options).getContentText();
  
  var document = XmlService.parse(responseMetadata);
  var root = document.getRootElement();
  var entityType = root.getChildren()[0].getChildren()[0].getChildren();
  for(var i =0;i<entityType.length;i++)
  {
    var elementName_ent = entityType[i].getName();
    if(elementName_ent == "EntityType")
    {
      var properties = entityType[i].getChildren();
      var tablename = entityType[i].getAttribute('Name').getValue();
      var table_properties = [];
      for(var j=0; j<properties.length; j++)
      {
        var elementName_prop = properties[j].getName();
        if(elementName_prop == "Property")
        {
          var propertyName = properties[j].getAttribute('Name').getValue();
          table_properties.push(propertyName);
        }
      }
      metadata[tablename] = table_properties;
    }
  }
  
  PropertiesService.getScriptProperties().setProperty('metadata', JSON.stringify(metadata));
  
  showTableSelectionDialog();
}

//Shows Table Selection Dialog
function showTableSelectionDialog(){
    var html = HtmlService.createTemplateFromFile('TableSelect');
    var metadata = JSON.parse(PropertiesService.getScriptProperties().getProperty('metadata'));
    html.data = Object.keys(metadata);
    var evaluated_html =  html.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setWidth(200).setHeight(100);
    SpreadsheetApp.getUi()
     .showModalDialog(evaluated_html, 'Select Table');
   
}

//Fetch Data from OData API after Table Selection
function fetchData(form)
{
  var options = {};
  
  var odataurl = PropertiesService.getScriptProperties().getProperty('odataurl');
  var username = PropertiesService.getScriptProperties().getProperty('username');
  var password = PropertiesService.getScriptProperties().getProperty('password');
  
  var table_name_selected = form.table_list;
  PropertiesService.getScriptProperties().setProperty('table_selected', table_name_selected);
  

  
  var odata_fetch_url = odataurl + table_name_selected + 'S';
  options.headers = {"Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)};
  var responseData = UrlFetchApp.fetch(odata_fetch_url, options).getContentText();
  var responseJSON = JSON.parse(responseData);
  
  //Set headers
  setHeaders();

  //SetData
  setData(responseJSON);
  
  if(responseJSON['@odata.nextLink'] != undefined){
    pageData(responseJSON['@odata.nextLink']);
  }
  
}

function pageData(url)
{
  var username = PropertiesService.getScriptProperties().getProperty('username');
  var password = PropertiesService.getScriptProperties().getProperty('password');
  
  var options = {};
  options.headers = {"Authorization": "Basic " + Utilities.base64Encode(username + ":" + password)};
  
  var responseData = UrlFetchApp.fetch(url, options).getContentText();
  var responseJSON = JSON.parse(responseData);
  
  setData(responseJSON);
  
   if(responseJSON['@odata.nextLink'] != undefined){
    pageData(responseJSON['@odata.nextLink']);
  }
  
}

//Set Sheet Headers
function setHeaders(){
  var table_name  = PropertiesService.getScriptProperties().getProperty('table_selected');
  var metadata = JSON.parse(PropertiesService.getScriptProperties().getProperty('metadata'));
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  sheet.clearContents();
  sheet.setFrozenRows(1);
  
  var table_columns = metadata[table_name];
  var endingRange = sheet.getRange(1, table_columns.length).getA1Notation();
  var range = "A1:" + endingRange;
  var totalRange = sheet.getRange(range);
  var headers = [table_columns];
  totalRange.setValues(headers);
}

//Append or Add Data to Sheet
function setData(responseJSON){
  
  var table_name  = PropertiesService.getScriptProperties().getProperty('table_selected');
  var metadata = JSON.parse(PropertiesService.getScriptProperties().getProperty('metadata'));
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheets()[0];
  var num_of_rows = responseJSON["value"].length;
  var records = responseJSON["value"];
  
  var table_columns = metadata[table_name];
  var num_of_columns = table_columns.length;
  var data_to_write = [];
  for(var i=0; i < num_of_rows; i++){
    var record_to_write = []
    for(var j=0; j < num_of_columns; j++)
    {
      record_to_write.push(records[i][table_columns[j]])
    }
    data_to_write.push(record_to_write);
  }
    sheet.getRange(sheet.getLastRow()+1, 1, data_to_write.length, data_to_write[0].length).setValues(data_to_write);
}
  

