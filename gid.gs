/* 머리행 공백, 2열에 #gid + 시트 ID
=ARRAYFORMULA(HYPERLINK(
 QUERY(INDEX(WORKSHEET(),,2), "offset 1"), 
 QUERY(INDEX(WORKSHEET(),,1), "offset 1")))
*/
function WORKSHEET() {
  try {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
    var out = new Array( sheets.length+1 ) ;
    out[0] = [ "", " " ];
    for (var i = 1 ; i < sheets.length+1 ; i++ ) out[i] = 
    [sheets[i-1].getName() , "#gid="+sheets[i-1].getSheetId() ];
    return out
  }
  catch( err ) {
    return "#ERROR!" 
    }
}
/* 머리행 있음, 2열에 시트 ID
=SHEETLIST()
=ARRAYFORMULA(IFNA(VLOOKUP(A1, HYPERLINK("#gid="&
 QUERY(INDEX(SHEETLIST();;2); "offset 1"); 
 QUERY(INDEX(SHEETLIST();;1); "offset 1")); 1; 0)))
*/
function SHEETLIST() {
  try {
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
    var out = new Array( sheets.length+1 ) ;
    out[0] = [ "워크시트" , "#gid" ];
    for (var i = 1 ; i < sheets.length+1 ; i++ ) out[i] = 
    [sheets[i-1].getName() , sheets[i-1].getSheetId() ];
    return out
  }
  catch( err ) {
    return "#ERROR!" 
    }
}
