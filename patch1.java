// Add "heading", "body copy", and "CTA button" for the "P2.1.1" row
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
    // Add "heading" in the same row as "P2.1.1"
    Row moduleRow = sheet.getRow(moduleRowIndex);
    Cell headingCell = moduleRow.createCell(masterElementsColIndex);
    headingCell.setCellValue("heading");

    // Shift rows below the "P2.1.1" row down by 2 positions to make room for "body copy" and "CTA button"
    sheet.shiftRows(moduleRowIndex + 1, sheet.getLastRowNum(), 2);

    // Add "body copy" in the row below
    Row bodyCopyRow = sheet.createRow(moduleRowIndex + 1);
    Cell bodyCopyCell = bodyCopyRow.createCell(masterElementsColIndex);
    bodyCopyCell.setCellValue("body copy");

    // Add "CTA button" in the row below
    Row ctaButtonRow = sheet.createRow(moduleRowIndex + 2);
    Cell ctaButtonCell = ctaButtonRow.createCell(masterElementsColIndex);
    ctaButtonCell.setCellValue("CTA button");
}