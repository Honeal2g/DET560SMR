/* Custom SMR Tracker (Updated 1 MAR 2020) * TSgt Ulan O Hawthorne Jr. * Detatchment 560 (Manhattan College) *
/* https://wingsuid.holmcenter.com/psc/wings_3/WINGS/WINGS_LOCAL/q/?ICAction=ICQryNameURL=PUBLIC.SMR_AS_FAVORITE*/
function onOpen() {
  var ui = SpreadsheetApp.getUi(), menu = ui.createMenu('SMR Options'), item = menu.addItem('Update Tracker','SMR_Code'); item.addToUi();
}
function SMR_Code(){
  if(DriveApp.getFilesByName("SMR_AS_FAVORITE.csv").hasNext() == true){
    var ss = SpreadsheetApp.getActiveSpreadsheet(), file = DriveApp.getFilesByName("SMR_AS_FAVORITE.csv").next(), csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
    DriveApp.getFilesByName('SMR_AS_FAVORITE.csv').next().setTrashed(true);    
    var dict = {}, Col_Index = [], CadetData = [];    
    var CustomColNames = ['EmplID','Name','AS Year','Stu Status','Major','SAT_COMP','AFOQT-Pilot','AFOQT-Nav','AFOQT-Apt','AFOQT-Verb','AFOQT-Quan','PCSM','Term GPA','Cum GPA','Term','Comm Dt','Enlist Date','Phys Exp','AFPFT','AFPFT Res','AFPFT Dt','MRS','Conditionals','Date of Birth','Citizen','Cat Sel','ACT-Score','Schlr Status','Schlr Type'];
    for(var i = 0; i < CustomColNames.length; i++){//Capture Col Indexes
      Col_Index[i] = csvData[0].indexOf(CustomColNames[i]);//stores index for respective column name   
    }
    for(var i = 1; i < csvData.length; i++){
      for(var j = 0; j < CustomColNames.length; j++){
        CadetData[j] = csvData[i][Col_Index[j]];
      }
      dict[Number(CadetData[0])]= CadetData; //Populate Dictionary
      CadetData = [];    
    }
    try{
      PushUpdates(dict);
      PushMajUpdates();
      CleanUpSheet();
      RefreshSheet();
      SortingFunction();
    }
    catch(error){
      Logger.log(error);
    }
    if(ss.getSheetByName('MajCode DB').isSheetHidden()){
      Logger.log("MajCode DB is Already Hidden");
    }else{
      ss.getSheetByName('MajCode DB').hideSheet();
    }      
    SpreadsheetApp.getActiveSpreadsheet().toast('Update Complete', 'Status', 5);
  }else{
    SortingFunction();    
    SpreadsheetApp.getActiveSpreadsheet().toast('SMR Already Up to Date', 'Status', 5);
  }
}
function MajorType(Flag){
  var Tech = {}, Non_Tech = {}, Col = ['F2:F','A2:A','B2:B','E2:E'], Data = ColData(Col,"MajCode DB");
  for(var i=0; i!=Data[1].length; i++){
    Tech[Data[1][i]] = Data[2][i];
    if(Data[3][i]!=""){Non_Tech[Data[3][i]] = Data[0][i];}
  }
  if(Flag == 1){return Tech;}
  if(Flag == 2){return Non_Tech;}
}
function SortingFunction() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getActiveSheet();
  spreadsheet.getRange('B1').activate();
  spreadsheet.getActiveSheet().getFilter().sort(2, true);
  spreadsheet.getRange('C1').activate();
  spreadsheet.getActiveSheet().getFilter().sort(3, false);  
}
function ColData(GrabCol,GrabSheet){
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet(), Sheet = activeSpreadsheet.getSheetByName(GrabSheet), Data = [], Ranges =[];
  for(var i=0; i!=GrabCol.length; i++){
    Temp = Sheet.getRange(GrabCol[i]);
    Ranges[i]= Temp;
  }  
  for(var i=0; i!=GrabCol.length; i++){
    Bin = Ranges[i].getValues();
    Data[i] = Bin;    
  }
  return Data;
}
function CleanUpSheet() {
  var spreadsheet = SpreadsheetApp.getActive(), sheet = spreadsheet.getActiveSheet();
  sheet.getRange(2, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  spreadsheet.getRange('5:5').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}
function RefreshSheet() {
  var spreadsheet = SpreadsheetApp.getActive(), sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).activate();
  spreadsheet.getActiveSheet().getFilter().remove();  
  sheet = spreadsheet.getActiveSheet();
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).createFilter();
  spreadsheet.getRange('E1').activate();
  var criteria = SpreadsheetApp.newFilterCriteria()
  .setHiddenValues(['Commissioned','Det Dropped'])
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(5, criteria);
}
function FindCols(ColNames) {
  var Col_Index = [], DetSMRCols = [], CustomColNames = ColNames, Internal_Data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SMR"), data = Internal_Data.getDataRange().getValues();  
  for(var i = 0; i < CustomColNames.length; i++){
    Col_Index[i] = data[0].indexOf(CustomColNames[i]);   
  }
  for(var i = 0; i < Col_Index.length; i++){
    var Range = Internal_Data.getRange(2,Col_Index[i]+1,data.length-1);
    DetSMRCols[i] = Range.getA1Notation();
  }
  return DetSMRCols;  
}
function PushMajUpdates(){
  var DetColRange = [], CadetRow = [], CustomColNames = ['Majcode','Major','Major Type'], DetCols = FindCols(CustomColNames), Tech_Dict = MajorType(1), Non_Tech_Dict = MajorType(2);
  var Internal_Data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SMR"), Maj_Type = Internal_Data.getRange(DetCols[2]), Major = Internal_Data.getRange(DetCols[1]);  
  for(var i = 0; i < DetCols.length; i++){
    DetColRange[i] = Internal_Data.getRange(DetCols[i]);
  }
  var DetColValue = DetColRange[0].getValues();
  for(var i=0; i!= DetColValue.length; i++){               
    if(Tech_Dict[DetColValue[i]]){
      Maj_Type.getCell(i+1, 1).setValue("Tech");
      Major.getCell(i+1, 1).setValue(Tech_Dict[DetColValue[i]]);     
    }
    if(Non_Tech_Dict[DetColValue[i]]){
      Maj_Type.getCell(i+1, 1).setValue("Non-Tech");
      Major.getCell(i+1, 1).setValue(Non_Tech_Dict[DetColValue[i]]);     
    }        
  }     
}  
function PushUpdates(Dictionary){
  var CustomColNames = ['EmpID','Name','AS-Level','Status','Majcode','SAT','Pilot','CSO/NAV','AA','Verbal','Quant','PCSM','TGPA','CGPA','TERM','DOC','DOE','DoDMERB-EXP','AFPT','AFPT-Stat','AFPT-DT','MRS','Conditionals','DOB','Citizen','Cat Sel','ACT','Schlr Status','Scholarship'];
  var DetCols = FindCols(CustomColNames), DetColRange = [], CadetRow = [], Internal_Data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SMR");  
  for(var i = 0; i < DetCols.length; i++){
    DetColRange[i] = Internal_Data.getRange(DetCols[i]);
  }
  var DetColValue = DetColRange[0].getValues();
  for(var i=0; i!= DetColValue.length; i++){               
    if(Dictionary[DetColValue[i]]){
      CadetRow = Dictionary[Number(DetColValue[i])]; 
      for(var j=1; j < CadetRow.length; j++){        
        if(CadetRow[j]){
          if(j == 2){//If we are updating AS-Level
            DetColRange[j].getCell(i+1, 1).setValue(Number(String(CadetRow[j]).substring(String(CadetRow[j]).indexOf("S")+1, String(CadetRow[j]).length))); //Removes the "AS" from "AS###" for easier spreadsheet sorting
          }else if(DetColRange[j]){                      
            DetColRange[j].getCell(i+1, 1).setValue(CadetRow[j]);
          }
        }
      }
    }        
  }     
}  
