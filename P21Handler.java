package com.example;

import org.apache.poi.ss.usermodel.*;

public class P21Handler {
    public static void handleP21(Workbook workbook, Workbook sourceWorkbook) {
        Sheet sheet = workbook.getSheetAt(0); // Assuming you are working with the first sheet.
        Sheet sourceSheet = sourceWorkbook.getSheet("EMAIL 1");

        int masterModulesColIndex = ExcelUtils.getColumnIndex(sheet, "Master_Modules");
        int sgEnContentColIndex = ExcelUtils.getColumnIndex(sheet, "SG-EN_Content");
        
        String moduleValue = "P2.1.1 - Left aligned primary copy module with background colour options";
        
        if (masterModulesColIndex != -1 && sgEnContentColIndex != -1) {
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row currentRow = sheet.getRow(i);
                if (currentRow != null) {
                    Cell masterModuleCell = currentRow.getCell(masterModulesColIndex);
                    if (masterModuleCell != null && masterModuleCell.getCellType() == CellType.STRING 
                        && masterModuleCell.getStringCellValue().equalsIgnoreCase(moduleValue)) {

                        // Create cells for heading, body copy, and CTA button
                        Cell headingCell = currentRow.createCell(sgEnContentColIndex);
                        headingCell.setCellValue("heading"); // Placeholder for heading

                        // Shift rows below the current row down by two positions
                        sheet.shiftRows(i + 1, sheet.getLastRowNum(), 2);

                        Row bodyCopyRow = sheet.createRow(i + 1);
                        bodyCopyRow.createCell(sgEnContentColIndex).setCellValue("bodycopy"); // Placeholder for body copy

                        Row ctaButtonRow = sheet.createRow(i + 2);
                        ctaButtonRow.createCell(sgEnContentColIndex).setCellValue("CTAbutton"); // Placeholder for CTA button

                        populateP21Content(sourceWorkbook, moduleValue, currentRow);
                        break; // Exit after processing the found module
                    }
                }
            }
        }
    }

    private static void populateP21Content(Workbook sourceWorkbook, String moduleValue, Row moduleRow) {
        Sheet sourceSheet = sourceWorkbook.getSheet("EMAIL 1");
        
        for (int i = 5; i <= sourceSheet.getLastRowNum(); i++) { 
            Row sourceDataRow = sourceSheet.getRow(i); 
            if (sourceDataRow != null) { 
                Cell moduleNameInSourceData = sourceDataRow.getCell(0); 

                if (moduleNameInSourceData != null && 
                    moduleNameInSourceData.getStringCellValue().equalsIgnoreCase(moduleValue)) { 

                    // Fetch content from E column for heading and body copy.
                    String headingContent = sourceDataRow.getCell(4) != null ? 
                        sourceDataRow.getCell(4).getStringCellValue() : ""; 
                    moduleRow.createCell(sgEnContentColIndex).setCellValue(headingContent); 

                    break; 
                } 
            } 
        } 
    }
}
