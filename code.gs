

function Refresh() {
 var spreadsheet = SpreadsheetApp.getActive();
SpreadsheetApp.flush()
}



function RefreshDate() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('18:18').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('@');


};

function color(range) {
return SpreadsheetApp.getActiveSheet().getRange(range).getBackground();
 }

function PushButon() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Targets'), true);
  RefreshDate();
  var spreadsheet = SpreadsheetApp.getActive();
};



function final2() {
  var spreadsheet = SpreadsheetApp.getActive();

/// targets
  spreadsheet.getRange('BJ40').activate();
 spreadsheet.getCurrentCell().setFormula('=if(BL40=true;Targets!B4;" ")')
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('BJ40:BJ52'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

 spreadsheet.getRange('BK40').activate();
 spreadsheet.getCurrentCell().setFormula('=if(BL40=true;HLOOKUP($BO$7;Targets!$B$64:$PJ$77;2);" ")')
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('BK40:BK52'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

 spreadsheet.getRange('BK41').activate();
 spreadsheet.getCurrentCell().setFormula('=if(BL41=true;HLOOKUP($BO$7;Targets!$B$64:$PJ$77;3);" ")')


 spreadsheet.getRange('BK42').activate();
 spreadsheet.getCurrentCell().setFormula('=if(BL42=true;HLOOKUP($BO$7;Targets!$B$64:$PJ$77;4);" ")')


 spreadsheet.getRange('BK43').activate();
 spreadsheet.getCurrentCell().setFormula('=if(BL43=true;HLOOKUP($BO$7;Targets!$B$64:$PJ$77;5);" ")')
 

 spreadsheet.getRange('BK44').activate();
 spreadsheet.getCurrentCell().setFormula('=if(BL44=true;HLOOKUP($BO$7;Targets!$B$64:$PJ$77;6);" ")')


spreadsheet.getRange('BK45').activate();
 spreadsheet.getCurrentCell().setFormula('=if(BL45=true;HLOOKUP($BO$7;Targets!$B$64:$PJ$77;7);" ")')


spreadsheet.getRange('BK46').activate();
 spreadsheet.getCurrentCell().setFormula('=if(BL46=true;HLOOKUP($BO$7;Targets!$B$64:$PJ$77;8);" ")')
 

 spreadsheet.getRange('BK47').activate();
 spreadsheet.getCurrentCell().setFormula('if(BL47=true;HLOOKUP($BO$7;Targets!$B$64:$PJ$77;9);" ")')
 

 spreadsheet.getRange('BK48').activate();
 spreadsheet.getCurrentCell().setFormula('=if(BL48=true;HLOOKUP($BO$7;Targets!$B$64:$PJ$77;10);" ")')


 spreadsheet.getRange('BK49').activate();
 spreadsheet.getCurrentCell().setFormula('=if(BL49=true;HLOOKUP($BO$7;Targets!$B$64:$PJ$77;11);" ")')


 spreadsheet.getRange('BK50').activate();
 spreadsheet.getCurrentCell().setFormula('=if(BL50=true;HLOOKUP($BO$7;Targets!$B$64:$PJ$77;12);" ")')


 spreadsheet.getRange('BK51').activate();
 spreadsheet.getCurrentCell().setFormula('=if(BL51=true;HLOOKUP($BO$7;Targets!$B$64:$PJ$77;13);" ")')


spreadsheet.getRange('BK52').activate();
 spreadsheet.getCurrentCell().setFormula('if(BL52=true;HLOOKUP($BO$7;Targets!$B$64:$PJ$77;14);" ")')

spreadsheet.getRange('BO8').activate();
 spreadsheet.getCurrentCell().setFormula('=IF(WEEKDAY(BO7;2)=1;0;IMPORTRANGE("1825wC22iyOJCs-FJgrj5TE3ioE2Wn69p75Oevoebgr8";CONCATENATE(CC5;TEXT( CE10;"dd-mm");CC5;CC6;CC8))+IMPORTRANGE("1825wC22iyOJCs-FJgrj5TE3ioE2Wn69p75Oevoebgr8";CONCATENATE(CC5;TEXT(CE10;"dd-mm");CC5;CC6;CC7)))')
spreadsheet.getRange('BO18').activate();
 spreadsheet.getCurrentCell().setFormula('=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1JnT_lRaXVK4jiOEvu4jGP8nLI8R3QCO7VqYBvnV42Zo/edit#gid=398193942";CONCATENATE(CC5;CD11;CC5;CC6;CE11))') 
spreadsheet.getRange('BO10').activate();
 spreadsheet.getCurrentCell().setFormula('=if(WEEKDAY(BO7;2)=1;IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"AJ204"))+IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"BK204"))+IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"CL204"))+IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"DM204"))+IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"EN204"))+IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"FO204"));if(WEEKDAY(BO7;2)=1;IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"BK204"))+IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"CL204"))+IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"DM204"))+IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"EN204"))+IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"FO204"));if(WEEKDAY(BO7;2)=3;IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"CL204"))+IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"DM204"))+IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"EN204"))+IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"FO204"));if(WEEKDAY(BO7;2)=4;IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"DM204"))+IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"EN204"))+IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"FO204"));if(WEEKDAY(BO7;2)=5;IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"EN204"))+IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"FO204"));if(WEEKDAY(BO7;2)=6;IMPORTRANGE("1EbQKyXxGRy7WkOxQ2DP9X40I14wVhAmU7J6yFN_eb9A";concatenate(Personnel!K2;VLOOKUP(WEEKNUM(BO7;1);Personnel!$I$1:$J$54;2;FALSE());Personnel!K2;Personnel!K3;"FO204"));1))))))')



  spreadsheet.getRange('D32').activate();
 spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(B32;Personnel!$A$1:$B$757;2;FALSE);"")')                                 
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('D32:D34'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('D7:D32'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('I33').activate();
 spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(G32;Personnel!$A$1:$B$757;2;FALSE);"")')
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('I33:I34'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('I7:I33'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('N33').activate();
 spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(L33;Personnel!$A$1:$B$757;2;FALSE);"")')
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('N33:N34'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('N7:N33'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('S33').activate();
 spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(Q33;Personnel!$A$1:$B$757;2;FALSE);"")')
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('S33:S34'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('S7:S33'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('X33').activate();
 spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(V33;Personnel!$A$1:$B$757;2;FALSE);"")')
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('X33:X34'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('X7:X33'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('AC33').activate();
 spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(aa33;Personnel!$A$1:$B$757;2;FALSE);"")')
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AC33:AC34'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AC7:AC33'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('AG33').activate();
 spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(AE33;Personnel!$A$1:$B$757;2;FALSE);"")')
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AG33:AG34'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AG7:AG33'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('AK33').activate();
 spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(AI33;Personnel!$A$1:$B$757;2;FALSE);"")')
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AK33:AK34'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AK7:AK33'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  spreadsheet.getRange('AO33').activate();
 spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(AM33;Personnel!$A$1:$B$757;2;FALSE);"")')
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AO33:AO34'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AO7:AO33'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('AS33').activate();
 spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(AQ33;Personnel!$A$1:$B$757;2;FALSE);"")')
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AS33:AS34'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AS7:AS33'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('AX33').activate();
 spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(AV33;Personnel!$A$1:$B$757;2;FALSE);"")')
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AX33:AX34'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('AX7:AX33'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('BC33').activate();
 spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(BA33;Personnel!$A$1:$B$757;2;FALSE);"")')
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('BC7:BC33'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('BC33:BC34'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('BH33').activate();
 spreadsheet.getCurrentCell().setFormula('=IFERROR(VLOOKUP(BF33;Personnel!$A$1:$B$757;2;FALSE);"")')
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('BH33:BH34'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('BH7:BH33'), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('BB36:BC36').activate();
  spreadsheet.getRange('BG36:BH37').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AW36:AX36').activate();
  spreadsheet.getRange('BG36:BH37').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AR36:AS36').activate();
  spreadsheet.getRange('BG36:BH37').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AN36:AO36').activate();
  spreadsheet.getRange('BG36:BH37').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AJ36:AK36').activate();
  spreadsheet.getRange('BG36:BH37').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AF36:AG36').activate();
  spreadsheet.getRange('BG36:BH37').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AB36:AC36').activate();
  spreadsheet.getRange('BG36:BH37').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('W36:X36').activate();
  spreadsheet.getRange('BG36:BH37').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('R36:S36').activate();
  spreadsheet.getRange('BG36:BH37').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('M36:N36').activate();
  spreadsheet.getRange('BG36:BH37').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('H36:I36').activate();
  spreadsheet.getRange('BG36:BH37').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('C36:D36').activate();
  spreadsheet.getRange('BG36:BH37').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('H6').activate();
  spreadsheet.getRange('C6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('M6').activate();
  spreadsheet.getRange('C6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('R6').activate();
  spreadsheet.getRange('C6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('W6').activate();
  spreadsheet.getRange('C6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AB6').activate();
  spreadsheet.getRange('C6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AF6').activate();
  spreadsheet.getRange('C6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AJ6').activate();
  spreadsheet.getRange('C6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AN6').activate();
  spreadsheet.getRange('C6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AR6').activate();
  spreadsheet.getRange('C6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('AW6').activate();
  spreadsheet.getRange('C6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('BB6').activate();
  spreadsheet.getRange('C6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('BG6').activate();
  spreadsheet.getRange('C6').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('B5:D5').activate();
  spreadsheet.getActiveRangeList().setBackground('#ff0000');
  spreadsheet.getRange('B7:D34').activate();
  spreadsheet.getActiveRangeList().setBackground('#ff0000');
  spreadsheet.getRange('B36:B37').activate();
  spreadsheet.getActiveRangeList().setBackground('#ff0000');
  spreadsheet.getRange('B5:D34').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID)
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('G7:I34').activate();
  spreadsheet.getActiveRangeList().setBackground('#6aa84f')
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('L7:N34').activate();
  spreadsheet.getActiveRangeList().setBackground('#a4c2f4')
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('L5:N5').activate();
  spreadsheet.getActiveRangeList().setBackground('#a4c2f4');
  spreadsheet.getRange('Q7:S34').activate();
  spreadsheet.getActiveRangeList().setBackground('#ff9900')
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('Q5:S5').activate();
  spreadsheet.getActiveRangeList().setBackground('#ff9900')
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('V7:X34').activate();
  spreadsheet.getActiveRangeList().setBackground('#d9d9d9')
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('V5:X5').activate();
  spreadsheet.getActiveRangeList().setBackground('#d9d9d9')
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AA5:AC5').activate();
  spreadsheet.getActiveRangeList().setBackground('#6aa84f');
  spreadsheet.getRange('AA7:AC34').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AE7:AG32').activate();
  spreadsheet.getActiveRangeList().setBackground('#ffd966')
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setBorder(false, false, false, false, false, false);
  spreadsheet.getRange('AE7:AG34').activate();
  spreadsheet.getActiveRangeList().setBorder(false, false, false, false, false, false)
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('B7:D34').activate();
  spreadsheet.getActiveRangeList().setBorder(false, false, false, false, false, false)
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('G7:I33').activate();
  spreadsheet.getActiveRangeList().setBorder(false, false, false, false, false, false)
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('G7:I34').activate();
  spreadsheet.getActiveRangeList().setBorder(false, false, false, false, false, false)
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('L7:N34').activate();
  spreadsheet.getActiveRangeList().setBorder(false, false, false, false, false, false)
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('Q7:S34').activate();
  spreadsheet.getActiveRangeList().setBorder(false, false, false, false, false, false)
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('V7:X34').activate();
  spreadsheet.getActiveRangeList().setBorder(false, false, false, false, false, false)
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AA7:AC34').activate();
  spreadsheet.getActiveRangeList().setBorder(false, false, false, false, false, false)
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AE7:AG34').activate();
  spreadsheet.getActiveRangeList().setBackground('#ffd966')
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setBackground('#ffd966');
  spreadsheet.getRange('AI7:AK34').activate();
  spreadsheet.getActiveRangeList().setBorder(false, false, false, false, false, false)
  .setBackground('#b6d7a8')
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AI5:AK5').activate();
  spreadsheet.getActiveRangeList().setBackground('#b6d7a8');
  spreadsheet.getRange('AM7:AO34').activate();
  spreadsheet.getActiveRangeList().setBackground('#ff9900')
  .setBorder(false, false, false, false, false, false)
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AQ7:AS34').activate();
  spreadsheet.getActiveRangeList().setBackground('#d9d9d9')
  .setBorder(false, false, false, false, false, false)
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AV7:AX34').activate();
  spreadsheet.getActiveRangeList().setBackground('#b6d7a8')
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AV5:AX5').activate();
  spreadsheet.getActiveRangeList().setBackground('#b6d7a8');
  spreadsheet.getRange('BA7:BC34').activate();
  spreadsheet.getActiveRangeList().setBackground('#b6d7a8')
  .setBorder(false, false, false, false, false, false)
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('BF7:BH34').activate();
  spreadsheet.getActiveRangeList().setBackground('#76a5af')
  .setBorder(false, false, false, false, false, false)
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('BA7:BC34').activate();
  spreadsheet.getActiveRangeList().setBorder(false, false, false, false, false, false)
  .setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('BA7:BC34').activate();
  spreadsheet.getActiveRangeList().setBackground('#b6d7a8');
  spreadsheet.getRange('BA30:BC34').activate();
  spreadsheet.getActiveRangeList().setBackground('#b6d7a8');
  spreadsheet.getRange('D2').activate();
  spreadsheet.getCurrentCell().setValue('La Sauce power');
  spreadsheet.getRange('B6:BH34').activate();
  spreadsheet.getActiveRangeList().setFontFamily('Arial')
  .setFontSize(10)
  .setFontWeight('bold')
  .setFontWeight(null)
  .setFontColor('#000000');
  spreadsheet.getRange('D2').activate();
  spreadsheet.getActiveRangeList().setFontSize(9);
  spreadsheet.getRange('G2').activate();
  
  spreadsheet.getRange('B5:D5').activate();
  spreadsheet.getCurrentCell().setFormula('=BJ40');
  spreadsheet.getRange('G5:I5').activate();
  spreadsheet.getCurrentCell().setFormula('=BJ41');
  spreadsheet.getRange('L5:N5').activate();
  spreadsheet.getCurrentCell().setFormula('=bj42');
  spreadsheet.getRange('Q5:S5').activate();
  spreadsheet.getCurrentCell().setFormula('=bj43');
  spreadsheet.getRange('V5:X5').activate();
  spreadsheet.getCurrentCell().setFormula('=bj44');
  spreadsheet.getRange('AA5:AC5').activate();
  spreadsheet.getCurrentCell().setFormula('=bj45');
  spreadsheet.getActiveRangeList().setFontColor('#000000');
  spreadsheet.getRange('AE5:AG5').activate();
  spreadsheet.getCurrentCell().setFormula('=bj46');
  spreadsheet.getRange('AI5:AK5').activate();
  spreadsheet.getCurrentCell().setFormula('=bj47');
  spreadsheet.getRange('AM5:AO5').activate();
  spreadsheet.getCurrentCell().setFormula('=bj48');
  spreadsheet.getRange('AQ5:AS5').activate();
  spreadsheet.getCurrentCell().setFormula('=bj49');
  spreadsheet.getRange('AV5:AX5').activate();
  spreadsheet.getCurrentCell().setFormula('=bj50');
  spreadsheet.getRange('BA5:BC5').activate();
  spreadsheet.getCurrentCell().setFormula('=bj51');
  spreadsheet.getRange('BF5:BH5').activate();
  spreadsheet.getCurrentCell().setFormula('=bj52');
  spreadsheet.getRange('AA5:AC5').activate();
  spreadsheet.getCurrentCell().setFormula('=bj45');
  spreadsheet.getRange('AA6').activate();

/// contours blancs
  spreadsheet.getRange('B5:D34').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('B36:D37').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('G36:I37').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('G5:I34').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('I34'));
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('L5:N34').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('L36:N37').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('Q36:S37').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('Q5:S34').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('Q34'));
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('V5:X34').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('V36:X37').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AA36:AC37').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AE36:AG37').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AI36:AK37').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AA5:AC34').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('AA34'));
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AE5:AG34').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AI5:AK34').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('AK34'));
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AM5:AO34').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AM36:AO37').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AQ36:AS37').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AQ5:AS34').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('AQ34'));
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AV5:AX34').activate();
  spreadsheet.getActiveRangeList().setBorder(false, false, false, false, false, false)
  .setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('AV36:AX37').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('BA36:BC37').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('BA5:BC34').activate();
  spreadsheet.setCurrentCell(spreadsheet.getRange('BC34'));
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('BF5:BH34').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('BF36:BH37').activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, null, null, '#f3f3f3', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  spreadsheet.getRange('BC29').activate();
  
  PushButon()
  var date = Utilities.formatDate(new Date, "UTC", "dd-MM");
  var source = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = source.getSheetByName(date);

  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(date), true);


};

// MAIL FOR KAROLINA


 function emailToKaputina() {

  // Send the PDF of the spreadsheet to this email address
  var email = "karolinaelizaokrasa@gmail.com";

  // Get the currently active spreadsheet URL (link)
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Subject of email message
  var subject = "Score report from the " + Utilities.formatDate(new Date(), "GMT", "dd-MMM-yyyy");

  // Email Body can  be HTML too
  var body = "Hello Kaputina, \n\nYou can find the scorereport of the day in the last page of the pdf. \n \nKind Regards\n \nLazare..";

  var blob = DriveApp.getFileById(ss.getId()).getAs("application/pdf");

  blob.setName(ss.getName() + ".pdf");

  // If allowed to send emails, send the email with the PDF attachment
  if (MailApp.getRemainingDailyQuota() > 0)
    GmailApp.sendEmail(email, subject, body, {
      htmlBody: body,
      attachments:[blob]
    });
}
///
function ConvertGoogleDocToCleanHtml() {
  var body = DocumentApp.getActiveDocument().getBody();
  var numChildren = body.getNumChildren();
  var output = [];
  var images = [];
  var listCounters = {};

  // Walk through all the child elements of the body.
  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    output.push(processItem(child, listCounters, images));
  }

  var html = output.join('\r');
  emailHtml(html, images);
  //createDocumentForHtml(html, images);
}

function emailHtml(html, images) {
  var attachments = [];
  for (var j=0; j<images.length; j++) {
    attachments.push( {
      "fileName": images[j].name,
      "mimeType": images[j].type,
      "content": images[j].blob.getBytes() } );
  }

  var inlineImages = {};
  for (var j=0; j<images.length; j++) {
    inlineImages[[images[j].name]] = images[j].blob;
  }

  var name = DocumentApp.getActiveDocument().getName()+".html";
  attachments.push({"fileName":name, "mimeType": "text/html", "content": html});
  MailApp.sendEmail({
     to: Session.getActiveUser().getEmail(),
     subject: name,
     htmlBody: html,
     inlineImages: inlineImages,
     attachments: attachments
   });
}

function createDocumentForHtml(html, images) {
  var name = DocumentApp.getActiveDocument().getName()+".html";
  var newDoc = DocumentApp.create(name);
  newDoc.getBody().setText(html);
  for(var j=0; j < images.length; j++)
    newDoc.getBody().appendImage(images[j].blob);
  newDoc.saveAndClose();
}

function dumpAttributes(atts) {
  // Log the paragraph attributes.
  for (var att in atts) {
    Logger.log(att + ":" + atts[att]);
  }
}

function processItem(item, listCounters, images) {
  var output = [];
  var prefix = "", suffix = "";

  if (item.getType() == DocumentApp.ElementType.PARAGRAPH) {
    switch (item.getHeading()) {
        // Add a # for each heading level. No break, so we accumulate the right number.
      case DocumentApp.ParagraphHeading.HEADING6: 
        prefix = "<h6>", suffix = "</h6>"; break;
      case DocumentApp.ParagraphHeading.HEADING5: 
        prefix = "<h5>", suffix = "</h5>"; break;
      case DocumentApp.ParagraphHeading.HEADING4:
        prefix = "<h4>", suffix = "</h4>"; break;
      case DocumentApp.ParagraphHeading.HEADING3:
        prefix = "<h3>", suffix = "</h3>"; break;
      case DocumentApp.ParagraphHeading.HEADING2:
        prefix = "<h2>", suffix = "</h2>"; break;
      case DocumentApp.ParagraphHeading.HEADING1:
        prefix = "<h1>", suffix = "</h1>"; break;
      default: 
        prefix = "<p>", suffix = "</p>";
    }

    if (item.getNumChildren() == 0)
      return "";
  }
  else if (item.getType() == DocumentApp.ElementType.INLINE_IMAGE)
  {
    processImage(item, images, output);
  }
  else if (item.getType()===DocumentApp.ElementType.LIST_ITEM) {
    var listItem = item;
    var gt = listItem.getGlyphType();
    var key = listItem.getListId() + '.' + listItem.getNestingLevel();
    var counter = listCounters[key] || 0;

    // First list item
    if ( counter == 0 ) {
      // Bullet list (<ul>):
      if (gt === DocumentApp.GlyphType.BULLET
          || gt === DocumentApp.GlyphType.HOLLOW_BULLET
          || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
        prefix = '<ul><li>', suffix = "</li>";

          suffix += "</ul>";
        }
      else {
        // Ordered list (<ol>):
        prefix = "<ol><li>", suffix = "</li>";
      }
    }
    else {
      prefix = "<li>";
      suffix = "</li>";
    }

    if (item.isAtDocumentEnd() || (item.getNextSibling() && (item.getNextSibling().getType() != DocumentApp.ElementType.LIST_ITEM))) {
      if (gt === DocumentApp.GlyphType.BULLET
          || gt === DocumentApp.GlyphType.HOLLOW_BULLET
          || gt === DocumentApp.GlyphType.SQUARE_BULLET) {
        suffix += "</ul>";
      }
      else {
        // Ordered list (<ol>):
        suffix += "</ol>";
      }

    }

    counter++;
    listCounters[key] = counter;
  }

  output.push(prefix);

  if (item.getType() == DocumentApp.ElementType.TEXT) {
    processText(item, output);
  }
  else {


    if (item.getNumChildren) {
      var numChildren = item.getNumChildren();

      // Walk through all the child elements of the doc.
      for (var i = 0; i < numChildren; i++) {
        var child = item.getChild(i);
        output.push(processItem(child, listCounters, images));
      }
    }

  }

  output.push(suffix);
  return output.join('');
}


function processText(item, output) {
  var text = item.getText();
  var indices = item.getTextAttributeIndices();

  if (indices.length <= 1) {
    // Assuming that a whole para fully italic is a quote
    if(item.isBold()) {
      output.push('<strong>' + text + '</strong>');
    }
    else if(item.isItalic()) {
      output.push('<blockquote>' + text + '</blockquote>');
    }
    else if (text.trim().indexOf('http://') == 0) {
      output.push('<a href="' + text + '" rel="nofollow">' + text + '</a>');
    }
    else {
      output.push(text);
    }
  }
  else {

    for (var i=0; i < indices.length; i ++) {
      var partAtts = item.getAttributes(indices[i]);
      var startPos = indices[i];
      var endPos = i+1 < indices.length ? indices[i+1]: text.length;
      var partText = text.substring(startPos, endPos);

      Logger.log(partText);

      if (partAtts.ITALIC) {
        output.push('<i>');
      }
      if (partAtts.BOLD) {
        output.push('<strong>');
      }
      if (partAtts.UNDERLINE) {
        output.push('<u>');
      }

      // If someone has written [xxx] and made this whole text some special font, like superscript
      // then treat it as a reference and make it superscript.
      // Unfortunately in Google Docs, there's no way to detect superscript
      if (partText.indexOf('[')==0 && partText[partText.length-1] == ']') {
        output.push('<sup>' + partText + '</sup>');
      }
      else if (partText.trim().indexOf('http://') == 0) {
        output.push('<a href="' + partText + '" rel="nofollow">' + partText + '</a>');
      }
      else {
        output.push(partText);
      }

      if (partAtts.ITALIC) {
        output.push('</i>');
      }
      if (partAtts.BOLD) {
        output.push('</strong>');
      }
      if (partAtts.UNDERLINE) {
        output.push('</u>');
      }

    }
  }
}


function processImage(item, images, output)
{
  images = images || [];
  var blob = item.getBlob();
  var contentType = blob.getContentType();
  var extension = "";
  if (/\/png$/.test(contentType)) {
    extension = ".png";
  } else if (/\/gif$/.test(contentType)) {
    extension = ".gif";
  } else if (/\/jpe?g$/.test(contentType)) {
    extension = ".jpg";
  } else {
    throw "Unsupported image type: "+contentType;
  }
  var imagePrefix = "Image_";
  var imageCounter = images.length;
  var name = imagePrefix + imageCounter + extension;
  imageCounter++;
  output.push('<img src="cid:'+name+'" />');
  images.push( {
    "blob": blob,
    "type": contentType,
    "name": name});
}
