function Main(){  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var wdcode = sheet.getRange("C2:C100").getValues();
  var status = sheet.getRange("E2:E100").getValues();
  var distNames = sheet.getRange("B2:B100").getValues();
  var rowNum = 0;
  var rowOfset = 2;
  var rowActual = rowNum+rowOfset;

  while (wdcode[rowNum][0]!=""){
   
    var managerEmail = FetchManagerMail(wdcode[rowNum][0]);
    
    //send mail
    var wdcodeName = distNames[rowNum][0];
    var  message = wdcodeName+" has filled the form please check row "+String(rowActual)+" in the Form !";
    var subject = wdcodeName+" Form filled !";
    console.log(managerEmail);
    MailApp.sendEmail(managerEmail, subject, message);

    //mark it done
    sheet.getRange("E"+String(rowActual)).setValue("Done");
    rowNum++;
    rowActual = rowNum+rowOfset;
  }

}




function FetchManagerMail(wdcode){
  //enter distributors code -> returns WD's manager email
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("WD AM Mapper");
  var wdcodes = sheet.getRange("A2:A100").getValues();

  var sheet3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AM Data");
  var managers = sheet3.getRange("A2:A100").getValues();
  var rowNum = 0;
  var rowOfset = 2;
  

  while(wdcodes[rowNum][0]!=wdcode){
    rowNum++;
  }
  var rowActual = rowNum+rowOfset;
  var managerCode = sheet.getRange("B"+String(rowActual)).getValue();
  
  var rowNum = 0;
  while(managers[rowNum][0]!=managerCode){
    rowNum++;
  }
  var rowActual = rowNum+rowOfset;
  var managerEmail = sheet3.getRange("C"+String(rowActual)).getValue();
  return managerEmail;
}



function sendManagerReport(managerCode){
  var nameEmail = amNameEmail(managerCode);
  var amName = nameEmail[0];
  var amEmail = nameEmail[1]; 
  console.log(amEmail);

  var managerWdCodeList = [];

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("WD AM Mapper");
  //sheet A-> WD Code    B-> AM Code   
  var managers = sheet.getRange("B2:B100").getValues();

  var rowNum = 0;
  while (managers[rowNum][0]!=""){
    if (managers[rowNum][0]==managerCode){
      var wdcode = sheet.getRange("A"+String(rowNum+2)).getValue();
      managerWdCodeList.push(wdcode);
    }
    rowNum++;
  }
  // console.log(managerWdCodeList);
  var wdCodeDone = [];
  var wdCodeNotDone = [];

  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  //sheet1 - > A:time   | B: name   | C: wd code   |  D: file   | E:  Status     
  
  var wdcodes = sheet1.getRange("C2:C100").getValues();
  var status = sheet1.getRange("E2:E100").getValues();
  var rowNum = 0;


  var notDoneCnt = 0;
  var doneCnt = 0;
  while (wdcodes[rowNum][0]!=""){
    
    if (managerWdCodeList.includes(wdcodes[rowNum][0])){
        wdCodeDone.push(wdcodes[rowNum][0]);
        doneCnt++;
    }
    rowNum++;
  }
  
  var totalWdCnt = 0;
  for (let code in managerWdCodeList){
    totalWdCnt++;
    if (!(wdCodeDone.includes(managerWdCodeList[code]))){
      wdCodeNotDone.push(managerWdCodeList[code]);
    }
  }
  notDoneCnt = totalWdCnt-doneCnt;
  
  // Message
  var subject = "OVI Report Status";
  var message ="Dear "+amName+"\n\n";
  message+="There are Currently "+notDoneCnt+" Distributors who have not sent the OVI Report and ";
  message+=doneCnt+" Distributors have filled the Report.\n\n";
  if(notDoneCnt>0){
    message+="Below is the List of Distributors who are yet to Send OVI Report : \n\n";
    for(let i in wdCodeNotDone){
      message+="Code: " + wdCodeNotDone[i]+"   Name: "+ wdName(wdCodeNotDone[i])+"\n";
    } 
  }
  if(doneCnt>0){
    message+="\nList of Distributors who have sent their OVI Report : \n\n";
    
    for(let i in wdCodeDone){
      message+="Code: "+ wdCodeDone[i]+"   Name: "+ wdName(wdCodeDone[i])+"\n";
    }
  }
  message+="\n\nThanks & Regards\nMohit Agarwal";
  MailApp.sendEmail(amEmail, subject, message);
  return message;
}
  



function updateManagers(){
  sheet =  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AM Data");
  managerCodes = sheet.getRange("A2:A100").getValues();
  var rowNum = 0;
  while(managerCodes[rowNum][0]!=""){
    console.log(managerCodes[rowNum][0]);
    //sendManagerReport(managerCodes[rowNum][0]);
    rowNum++;
  }
 
}


function wdName(wdcode){
  //returns wd name takes in wdcode
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("WD Data");
  code = sheet.getRange("A2:A100").getValues();

  rowNum = 0;
  while(code[rowNum][0]!=wdcode){
    rowNum++;
  }
  return sheet.getRange("B"+String(rowNum+2)).getValue();
  
}


function amNameEmail(amcode){
  //takes in amcode returns name and mail of Manager

  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("AM Data");
  //AM DAta
  // A: AM Code -> B: AM Name -> C : AM Email
  codes = sheet.getRange("A2:A100").getValues();

  var rowNum = 0;
  while (codes[rowNum][0]!=amcode){
    rowNum++;
  }

  var mail  = sheet.getRange("C"+String(rowNum+2)).getValue();
  var name = sheet.getRange("B"+String(rowNum+2)).getValue();
  return [name,mail];

}






