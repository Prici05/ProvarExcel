// Find the corresponding row in Email 1 sheet for the Master Module value
String masterModuleValue = "P2.1.1 - Left aligned primary copy module with background colour options"; // Example value
int masterModuleRowIndex = -1;
for (int i = 5; i <= sourceSheet.getLastRowNum(); i++) {
    Row sourceRow = sourceSheet.getRow(i);
    if (sourceRow != null) {
        Cell moduleCell = sourceRow.getCell(reqcolindex);
        if (moduleCell != null && moduleCell.getCellType() == CellType.STRING &&
                moduleCell.getStringCellValue().equalsIgnoreCase(masterModuleValue)) {
            masterModuleRowIndex = sourceRow.getRowNum();
            break;
        }
    }
}

// If row is found, fetch content from E column (column index 4) and populate SG_EN Content
if (masterModuleRowIndex != -1) {
    // Fetch content for heading (current row) and body copy (next row) from column E
    Row sourceRow = sourceSheet.getRow(masterModuleRowIndex);
    Cell contentCellHeading = sourceRow.getCell(4); // Column E (index 4) for heading
    String headingContent = contentCellHeading != null && contentCellHeading.getCellType() == CellType.STRING
            ? contentCellHeading.getStringCellValue()
            : "";

    Row nextRow = sourceSheet.getRow(masterModuleRowIndex + 1);
    Cell contentCellBody = nextRow != null ? nextRow.getCell(4) : null; // Column E for body copy
    String bodyCopyContent = contentCellBody != null && contentCellBody.getCellType() == CellType.STRING
            ? contentCellBody.getStringCellValue()
            : "";

    // Populate SG_EN Content in ProvarExcel for Heading and Body Copy rows
    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
        Row currentRow = sheet.getRow(i);
        if (currentRow != null) {
            Cell masterModuleCell = currentRow.getCell(masterModulesColIndex);
            if (masterModuleCell != null && masterModuleCell.getCellType() == CellType.STRING &&
                    masterModuleCell.getStringCellValue().equalsIgnoreCase(masterModuleValue)) {
                // Set SG_EN Content for the Heading Row
                Cell sgEnContentCell = currentRow.createCell(sgEnContentColIndex);
                sgEnContentCell.setCellValue(headingContent);

                // Now, set SG_EN Content for the Body Copy row (next row)
                Row bodyCopyRow = sheet.getRow(i + 1); // Move to next row for body copy
                if (bodyCopyRow != null) {
                    Cell sgEnBodyCopyCell = bodyCopyRow.createCell(sgEnContentColIndex);
                    sgEnBodyCopyCell.setCellValue(bodyCopyContent);
                }
                break;
            }
        }
    }
}