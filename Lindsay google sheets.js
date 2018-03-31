function reflectCell (tgt, sFrom, sWkTo, sCellTo, colour) {
  // all the inputs have to have values
  if (tgt != null && sWkTo != null && sFrom != null && sCellTo != null) {
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    
    // get the ranges for the from and to cells
    ssFrom = ss.getSheetByName(tgt);
    cellFrom = ssFrom.getRange(sFrom);
    ssTo = ss.getSheetByName(sWkTo);
    cellTo = ssTo.getRange(sCellTo)
    // do the copy
    cellFrom.copyTo(cellTo);
                
    // this bit copies the color as well based on what is passed to col - defaults to true so you only need to
    // set it if you dont want perty colors.
    if (colour == true) {
      cellTo.setBackground(cellFrom.getBackground());
    }
  }
}

function onChange(e) {
  // there is probably a way to filter the events to the copy isn't done unless the onChange cell is the one we are interested in
  // but I can't be bothered figuring it out and this works - that improvement is left as an exercise to the user.
  reflectCell ("TakeOff", "A2", "Legend", "E5", true);
  reflectCell ("TakeOff", "A2", "BD", "B13", true);
}
