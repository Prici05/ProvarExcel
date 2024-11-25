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
    voCell.setCellValue("// Add rows for "heading", "body copy", and "CTA button" for the "P2.1.1" row
int moduleRowIndex = -1;
String moduleValue = "P2.1.1 - Left aligned primary copy module with background colour options";

// Find the row with the "Master_Modules" value "P2.1.1 - Left aligned primary copy module with background colour options"
for (int i = 1; i <= sheet.getLastRowNum(); i++) {
    Row currentRow = sheet.getRow(i);
    if (currentRow != null) {
        Cell moduleCell = currentRow.getCell(masterModulesColIndex);
        if (moduleCell != null && moduleCell.getCellType() == CellType.STRING &&
                moduleCell.getStringCellValue().equalsIgnoreCase(moduleValue)) {
            moduleRowIndex = currentRow.getRowNum();
            break;
        }
    }
}

if (moduleRowIndex != -1 && masterElementsColIndex != -1) {
    // Shift rows below the "P2.1.1" row down by 3 positions to make room for the new rows
    sheet.shiftRows(moduleRowIndex + 1, sheet.getLastRowNum(), 3);

    // Add "heading", "body copy", and "CTA button" in the rows below
    Row headingRow = sheet.createRow(moduleRowIndex + 1);
    Cell headingCell = headingRow.createCell(masterElementsColIndex);
    headingCell.setCellValue("heading");

    Row bodyCopyRow = sheet.createRow(moduleRowIndex + 2);
    Cell bodyCopyCell = bodyCopyRow.createCell(masterElementsColIndex);
    bodyCopyCell.setCellValue("body copy");

    Row ctaButtonRow = sheet.createRow(moduleRowIndex + 3);
    Cell ctaButtonCell = ctaButtonRow.createCell(masterElementsColIndex);
    ctaButtonCell.setCellValue("CTA button");
}


