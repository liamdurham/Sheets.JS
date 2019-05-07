var roster = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster');
var input = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Input');
var blacklist = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Blacklist');
var pd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('P/D');
var discharge = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Discharge');
var act = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Activity');
var loa = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('LOA/ROA');
var jr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Jedi Roster');
var tlog = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('TL');
var subd = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Subdivision Input');
var pdl = pd.getRange("A:G");
var rl = roster.getRange("A:T");
var tl = tlog.getRange("A:H");
var rcl = input.getRange("A:I");
var dscg = discharge.getRange("A:G");
var abs = loa.getRange("A:I");
var jedi = jr.getRange("A:Q");
var sd = subd.getRange("A:I");
var sdr = subd.getDataRange().getDisplayValues();
var range = roster.getDataRange().getDisplayValues();
var work = input.getDataRange().getDisplayValues();
var name = "";
var rank = "";
var activecol = "0";
var activerow = "0";
var newrank = "";
var aspot = "";
var rankArray = ["PVT","PFC","LCPL","CPL","SGT","SSG","SFC","MSG","1SG","SGM","CSM","SMB","2ndLT","1stLT","CPT","MAJ","LTC","COL","XO","CMD"];
var seenArray = [];
var id = "";

function onEdit() {
  Logger.log((" edit 1" )); 
 rl.sort({column: 1, ascending: false});  
  Logger.log((" edit 2" )); 
 rcl.sort({column: 1, ascending: false});  
  Logger.log((" edit 3" )); 
 pdl.sort({column: 1, ascending: false}); 
  Logger.log((" edit 4" )); 
 dscg.sort({column: 1, ascending: false}); 
  Logger.log((" edit 5" )); 
 jedi.sort({column: 1, ascending: false}); 
  Logger.log((" edit 6" )); 
 abs.sort({column: 2, ascending: false});  
  Logger.log((" edit 7" )); 
 sd.sort({column: 1, ascending: false});
  Logger.log((" edit 8" )); 
 tl.sort({column: 1, ascending: false});
  Logger.log((" edit 9" )); 
}

function onOpen(){  
  sdr.forEach(function(row1) {
    if (row1[3] != ""){
   Logger.log(("Working on "+activerow+" - "+row1[3]+"."));
   activerow++;
    if (row1[3] == "Lead")
   {
     var a = activerow;
     var c = row1[4];
     Logger.log((" ---- " +seenArray)); 
     Logger.log((" Seen " +row1[4]+ "."));
     if(seenArray.indexOf(c) != -1)
     {
      Logger.log((" !!! Deleteing " +row1[3]));
      for(var r=1;r<7;r++){
      var a = subd.getRange(activerow,r);
      a.setValue("");   
     }
     }
    seenArray.push(row1[4]); 
   }}
  });
 
 activerow = 0;
 var today = new Date();
 var week = new Date(new Date()-(60*60*24*8*1000) );
 Logger.log((week)); 
 range.forEach(function(row){ 
    name = row[3];
    rank = row[1];
    id = row[5];
    joinedDate = new Date(row[15]);
    if (row[3] != "")
    {
    Logger.log((name+"--"+rank+"--"+week+">="+joinedDate+"--")); 
   if(rankArray.indexOf(rank) < 4 ){ 
     if(week > joinedDate){
       if(name != "Name"){
     Logger.log(("---"+name+" Marked for deletion."));
     discharge.appendRow([today,Session.getActiveUser().getEmail(),"N/A",name,rank,"Automatic 1wk inactive enlisted discharge",id])
     act.appendRow([today,Session.getActiveUser().getEmail(),"AutoDischarged",name]);
     work.forEach(function(rows) {
     activecol ++;
       if (rows[5] == name) {  
         for(var r=6;r<9;r++){
           Logger.log((name+" Deleted."));
           var a = input.getRange(activecol,r);
           a.setValue(""); 
          }}});  
        activecol = 0;
        r = 0;
        Logger.log(("AutoDischarge "+name+" found and complete."));
       }}}
    }
    }); 
  onEdit();
}

function onButton()
{   
  activerow = 0;
  Logger.log(("Start"));
  range.forEach(function(row){
   activerow ++;
   if (row[17] != "Reset")
   {
    if (row[17] != "")
    {
   switch(row[17]) {
        case "Promote":
            name = row[3];
            rank = row[1];
            aspot = rankArray.indexOf(rank);
            Logger.log(("Promo "+name+" from "+rank+" requested."));
                  if(rank != "CMD") {                   
                    newrank = rankArray.slice(aspot+1,aspot+2).toString();
                    Logger.log(("Promo "+name+" from "+rank+" to "+newrank+"."));
                    var today = new Date();
                    pd.appendRow([today,Session.getActiveUser().getEmail(),"N/A","Promotion",name,newrank,"Officer Manual Rankup"]);
                    act.appendRow([today,Session.getActiveUser().getEmail(),"Promoted",name]);
                   roster.getRange(activerow, 18).setValue("Reset");
                  }
            break;
        case "Demote":  
            name = row[3];
            rank = row[1];
            aspot = rankArray.indexOf(rank);
            Logger.log(("Demote "+name+" from "+rank+" requested." +aspot+ "."));
                  if(rank != "PVT") {                   
                    newrank = rankArray.slice(aspot-1,aspot).toString();
                    Logger.log(("Demote "+name+" from "+rank+" to "+newrank+"."));
                    var today = new Date();
                    pd.appendRow([today,Session.getActiveUser().getEmail(),"N/A","Demotion",name,newrank,"Officer Manual Demotion"]);
                    act.appendRow([today,Session.getActiveUser().getEmail(),"Demoted",name]);
                   roster.getRange(activerow, 18).setValue("Reset");
                  }
            break;
       case "Discharge":
            name = row[3];
            rank = row[1];
            id = row[5];
                    var today = new Date();
                    discharge.appendRow([today,Session.getActiveUser().getEmail(),"N/A",name,rank,"Officer Mass Discharge",id])
                    act.appendRow([today,Session.getActiveUser().getEmail(),"Discharged",name]);
                    work.forEach(function(rows) {
                    activecol ++;
                    if (rows[5] == name) {  
                    for(var r=6;r<9;r++){
                    var a = input.getRange(activecol,r);
                    a.setValue(""); 
                     }}});  
                    activecol = 0;
                    r = 0;
                    roster.getRange(activerow, 18).setValue("Reset");
                    Logger.log(("Discharge "+name+" found and complete."));            
            break;
        case "Blacklist 2wk":
            name = row[3];
            rank = row[1];
          var id = row[5];
                    var today = new Date();
                    var end = new Date(Date.now() + 12096e5);
                    blacklist.appendRow([today,Session.getActiveUser().getEmail(),end,id,name,rank,"Officer Manual Blacklist"])
                    act.appendRow([today,Session.getActiveUser().getEmail(),"Blacklisted 2wks",name]);
                    work.forEach(function(rows) {
                    activecol ++;
                    if (rows[5] == name) {  
                    for(var r=6;r<9;r++){
                    var a = input.getRange(activecol,r);
                    a.setValue(""); 
                     }}});  
                    activecol = 0;
                    r = 0;
                    roster.getRange(activerow, 18).setValue("Reset");
                    Logger.log(("Blacklist 2wks "+name+" found and complete."));   
           break;  
        case "Blacklist Perm":
             name = row[3];
             rank = row[1];
           var id = row[5];
                    var today = new Date();
                    blacklist.appendRow([today,Session.getActiveUser().getEmail(),"PERMANENT",id,name,rank,"Officer Manual Blacklist"])
                    act.appendRow([today,Session.getActiveUser().getEmail(),"Blacklisted PERM",name]);
                    work.forEach(function(rows) {
                    activecol ++;
                    if (rows[5] == name) {  
                    for(var r=6;r<9;r++){
                    var a = input.getRange(activecol,r);
                    a.setValue(""); 
                     }}});  
                    activecol = 0;
                    r = 0;
                    roster.getRange(activerow, 18).setValue("Reset");
                    Logger.log(("Blacklist Perm "+name+" found and complete."));   
           break;  
         default:
           break; 
   }}}}); 
  onEdit();
}

      