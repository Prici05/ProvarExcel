// Add "ps", "ssl", and "vo" under "Master_Elements" for the "Pre Header" section
if (preHeaderRowIndex != -1 && masterElementsColIndex != -1) {
    // Add "ps" in the same row as "Pre Header"
    Row preHeaderRow = sheet.getRow(preHeaderRowIndex);
    Cell psCell = preHeaderRow.createCell(masterElementsColIndex);
    psCell.setCellValue("ps");

    // Shift rows below the "Pre Header" row down by 2 positions to make room for "ssl" and "vo"
    sheet.shiftRows(preHeaderRowIndex + 1, sheet.getLastRowNum(), 2);

    // Add "ssl" and "vo" in the rows below
    Row sslRow = sheet.createRow(preHeaderRowIndex + 1);
    Cell sslCell = sslRow.createCell(masterElementsColIndex);
    sslCell.setCellValue("ssl");

    Row voRow = sheet.createRow(preHeaderRowIndex + 2);
    Cell voCell = voRow.createCell(masterElementsColIndex);
    voCell.setCellValue("vo");
}
