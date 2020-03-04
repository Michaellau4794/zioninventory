//https://www.youtube.com/playlist?list=PLv9Pf9aNgemt82hBENyneRyHnD-zORB3l

var url = "https://docs.google.com/spreadsheets/d/1mYUxm6zX8Gs0K594m5FiyUGXSUzwgUWx_y1_4t_AjXw/edit#gid=0";
var beurl = "https://docs.google.com/spreadsheets/d/1RV28QjrRuOhQ1NT6ivLPljzy_yu8o-SBrxixYUnH_Pg/edit#gid=141623591";

function doGet(e) {

  if(e.parameters.v == "form"){
  return loadForm();
  } else if(e.parameters.v == "enquiry"){
      return loadEnquiry2();
  }
   return HtmlService.createTemplateFromFile('home').evaluate().setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }




function loadForm(){

   var ss = SpreadsheetApp.openByUrl(url);
   var ws = ss.getSheetByName("Options");
   var list = ws.getRange(1, 1,ws.getRange("A1").getDataRegion().getLastRow(),1).getValues();
  var list2 = ws.getRange(1, 2,ws.getRange("B1").getDataRegion().getLastRow(),1).getValues();
   var htmlListArray = list.map(function(r){ return '<option>' + r[0] + '</option>'; }).join('');
  var htmlListArray2 = list2.map(function(t){ return '<option>' + t[0] + '</option>'; }).join('');

  var tmp = HtmlService.createTemplateFromFile('page');
  tmp.list = htmlListArray;
  tmp.list2 = htmlListArray2;

  return tmp.evaluate()
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


///////////////////////////////////////////////////////
//function loadEnquiry(){
//
//   var ss = SpreadsheetApp.openByUrl(url);
//  var bess = SpreadsheetApp.openByUrl(url);
//   var ws = ss.getSheetByName("Options");
//
////  Mutliple Choise Option
//   var list = ws.getRange(1, 1,ws.getRange("A1").getDataRegion().getLastRow(),1).getValues();
//  var list2 = ws.getRange(1, 3,ws.getRange("C1").getDataRegion().getLastRow(),1).getValues();
//   var htmlListArray = list.map(function(r){ return '<option>' + r[0] + '</option>'; }).join('');
//  var htmlListArray2 = list2.map(function(t){ return '<option>' + t[0] + '</option>'; }).join('');
//  Logger.log(ws.getRange("A1").getDataRegion().getLastRow());
//  Logger.log(ws.getRange("C1").getDataRegion().getLastRow());
//
////  Table Content -
//   var wsData = ss.getSheetByName("CalData");
//   var Tlist = wsData.getRange(1, 1,wsData.getRange("A1").getDataRegion().getLastRow(),6).getValues();
//  Logger.log("Tlist: " + Tlist);
//  var rNum = wsData.getRange("A1").getDataRegion().getLastRow();
//  var TArray = [];
//
//  for (let i=1; i<rNum; i++){
//    if(Tlist[i][4] == ""){
//
//    } else if ( Tlist[i][0] == "新增倉存" || Tlist[i][0] == "減少倉存" ) {
//      var TListArray = '<tr>' + '<td>' + Tlist[i][0] + '</td>' + '<td>' + Tlist[i][1] + '</td>' + '<td>' + Tlist[i][2] + '</td>' + '<td>' + Tlist[i][3] + '</td>'  + '<td>' + (Tlist[i][4]).toLocaleDateString() + '</td>' + '</tr>';
//      TArray.push(TListArray);
//    }
//  }
//
//
//  var TArr = TArray.join('');
//  Logger.log(TArr);
//
//  var tmp = HtmlService.createTemplateFromFile('enquiry');
//  tmp.list = htmlListArray;
//  tmp.list2 = htmlListArray2;
//  tmp.Tlist = TArr;
//
//  return tmp.evaluate()
//  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
//}


///////////////////////////////////////////////////////
function loadEnquiry2(){

   var ss = SpreadsheetApp.openByUrl(url);
  var bess = SpreadsheetApp.openByUrl(beurl);
   var ws = bess.getSheetByName("Option");

//  Mutliple Choise Option
   var list = ws.getRange(1, 1,ws.getRange("A1").getDataRegion().getLastRow(),1).getValues();
  var list2 = ws.getRange(1, 3,ws.getRange("C1").getDataRegion().getLastRow(),1).getValues();
   var htmlListArray = list.map(function(r){ return '<option>' + r[0] + '</option>'; }).join('');
  var htmlListArray2 = list2.map(function(t){ return '<option>' + t[0] + '</option>'; }).join('');
  Logger.log(ws.getRange("A1").getDataRegion().getLastRow());
  Logger.log(ws.getRange("C1").getDataRegion().getLastRow());

//  Table Content - All Txn
   var wsData = bess.getSheetByName("CalculatedData");
   var Tlist = wsData.getRange(1, 1,wsData.getRange("A1").getDataRegion().getLastRow(),6).getValues();
  Logger.log("Tlist: " + Tlist);
  var rNum = wsData.getRange("A1").getDataRegion().getLastRow();
  var TArray = [];

  for (let i=1; i<rNum; i++){
    if(Tlist[i][4] == ""){

    } else if ( Tlist[i][1] == "新增倉存" || Tlist[i][1] == "減少倉存" ) {
      var TListArray = '<tr>' + '<td>' + (Tlist[i][0]).toLocaleDateString() +
          '</td>' + '<td>' + Tlist[i][1] +
            '</td>' + '<td>' + Tlist[i][2] +
              '</td>' + '<td>' + Tlist[i][3] +
                '</td>' + '<td>' + Tlist[i][4] +
                  '</td>'  + '<td>' + (Tlist[i][5]).toLocaleDateString() + '</td>' + '</tr>';
      TArray.push(TListArray);
    }
  }

// Table Content - Summary
  var wsData2 = bess.getSheetByName("CalculatedDataPTable");
  var Tlist2 = wsData2.getRange(1, 1,wsData2.getRange("A1").getDataRegion().getLastRow(),3).getValues();
  Logger.log("Tlist2: " + Tlist2);
  var rNum2 = wsData2.getRange("A1").getDataRegion().getLastRow();
  var TArray2 = [];

  for (let i=1; i<rNum2; i++){
    if(Tlist2[i][0] == "" || Tlist2[i][0] == "地點" ){

    } else {
      var TListArray2 = '<tr>' + '<td>' + (Tlist2[i][0]) +
          '</td>' + '<td>' + Tlist2[i][1] +
            '</td>' + '<td>' + Tlist2[i][2] + '</td>' + '</tr>';
      TArray2.push(TListArray2);
    }
  }


  var TArr = TArray.join('');
  Logger.log(TArr);
  var TArr2 = TArray2.join('');
  Logger.log(TArr);

  var tmp = HtmlService.createTemplateFromFile('enquiry');
  tmp.list = htmlListArray;
  tmp.list2 = htmlListArray2;
  tmp.Tlist = TArr;
  tmp.Tlist2 = TArr2;



  return tmp.evaluate()
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
