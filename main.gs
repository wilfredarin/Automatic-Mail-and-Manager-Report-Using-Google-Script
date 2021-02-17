//enter form Response Sheet Name
var responseSheetName = "Form Responses 1";

//enter WD AM Mapper Sheet Name
var wdAmMapperSheet = "WD AM Mapper";

//enter AM Data Sheet Name
var amDataSheet = "AM Data";

//enter WD Data Sheet Name
var wdDataSheet = "WD Data";


function main(){  

  //Sends Mail to the Assigned AM for every File Upload By WD 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(responseSheetName);
  //sheet : A:Time --> B:Name --> C: WD Code  --> D: File --> E: Status
  var wdcodes = sheet.getRange("C2:C100").getValues();
  var status = sheet.getRange("E2:E100").getValues();
  var wdNames = sheet.getRange("B2:B100").getValues();
  
  var rowNum = 0;

  // check form enteries untill empty row found
  while (wdcodes[rowNum][0]!=""){
    //if status is not done Send The Mail
    if (status[rowNum][0]!="Done"){
      var managerCode = fetchManagerCode(wdcodes[rowNum][0]);
      var nameEmail = fetchAmNameEmail(managerCode);
      var managerName = nameEmail[0];
      var managerEmail = nameEmail[1];
      
      //send mail
      var wdcodeName = wdNames[rowNum][0];
      var subject = wdcodeName+" has sent OVI Report !";
      var message = "Dear "+ managerName+"\n\n";
      message += wdcodeName +", WD Code  "+wdcodes[rowNum][0];
      message +=" has filled the form please check row "+String(rowNum+2)+" in the Form !\n";
      message +="\n\n\nThanks & Regards\nMohit Agarwal"
    
      MailApp.sendEmail(managerEmail, subject, message);

      //mark it done
      sheet.getRange("E"+String(rowNum+2)).setValue("Done");

      //break assuming all before it are done and every form filled runs this function
      break;
    }

    rowNum++;
  }

}




function fetchManagerCode(wdcode){
  //enter distributors code -> returns WD's manager email
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(wdAmMapperSheet);
  // A: WD CODE --> B : AM CODE 
  var wdcodes = sheet.getRange("A2:A100").getValues();
  var rowNum = 0;
  while(wdcodes[rowNum][0]!=wdcode){
    rowNum++;
  }  
  var managerCode = sheet.getRange("B"+String(rowNum+2)).getValue();
  return managerCode;
}




function sendManagerReport(managerCode){
  //to send report to manager --> enter manager Code
  var nameEmail = fetchAmNameEmail(managerCode);
  var amName = nameEmail[0];
  var amEmail = nameEmail[1]; 
  var managerWdCodeList = [];
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(wdAmMapperSheet);
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
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(responseSheetName);
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
  // *this needs to updated later for the case when all have sent and none have sent*
  var message ="Dear "+amName+"\n\n";
  message+="There are Currently "+notDoneCnt+" Distributors who have not sent the OVI Report and ";
  message+=doneCnt+" Distributors have filled the Report.\n\n";
  if(notDoneCnt>0){
    message+="Below is the List of Distributors who are yet to Send OVI Report : \n\n";
    for(let i in wdCodeNotDone){
      message+="Code: " + wdCodeNotDone[i]+"   Name: "+ fetchWdName(wdCodeNotDone[i])+"\n";
    } 
  }
  if(doneCnt>0){
    message+="\nList of Distributors who have sent their OVI Report : \n\n";
    
    for(let i in wdCodeDone){
      message+="Code: "+ wdCodeDone[i]+"   Name: "+ fetchWdName(wdCodeDone[i])+"\n";
    }
  }
  message+="\n\nThanks & Regards\nMohit Agarwal";
  return message;
}
  






function fetchWdName(wdcode){
  //returns wd name takes in wdcode
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(wdDataSheet);
  code = sheet.getRange("A2:A100").getValues();
  rowNum = 0;
  while(code[rowNum][0]!=wdcode){
    rowNum++;
  }
  return sheet.getRange("B"+String(rowNum+2)).getValue();
}


function fetchAmNameEmail(amcode){
  //takes in amcode returns name and mail of Manager
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(amDataSheet);
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

