package com.example;

import org.apache.poi.ss.usermodel.*;

public class P21Handler {
    public static void handleP21(Workbook workbook, Workbook sourceWorkbook) {
        Sheet sheet = workbook.getSheetAt(0); // Assuming you are working with the first sheet.
        Sheet sourceSheet = sourceWorkbook.getSheet("EMAIL 1");

        int masterModulesColIndex = ExcelUtils.getColumnIndex(sheet, "Master_Modules");
        int masterElementsColIndex = ExcelUtils.getColumnIndex(sheet, "Master_Elements");
        
        String moduleValue = "P2.1.1 - Left aligned primary copy module with background colour options";
        
        int moduleRowIndex = findModuleByName(sheet, moduleValue);

        if (moduleRowIndex != -1 && masterElementsColIndex != -1) {
            Row moduleRow = sheet.getRow(moduleRowIndex);
            Cell headingCell = moduleRow.createCell(masterElementsColIndex);
            headingCell.setCellValue("heading");

            // Shift rows below the P2.1.1 row down by two positions to make room for body copy and CTA button.
            sheet.shiftRows(moduleRowIndex + 1, sheet.getLastRowNum(), 2);

            Row bodyCopyRow = sheet.createRow(moduleRowIndex + 1);
            bodyCopyRow.createCell(masterElementsColIndex).setCellValue("bodycopy");

            Row ctaButtonRow = sheet.createRow(moduleRowIndex + 2);
            ctaButtonRow.createCell(masterElementsColIndex).setCellValue("CTAbutton");

            populateP21Content(sourceWorkbook, moduleValue, moduleRow);
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
                    moduleRow.createCell(masterElementsColIndex).setCellValue(headingContent); 

                    break; 
                } 
            } 
        } 
    }

    private static int findModuleByName(Sheet sheet, String value) {
        for (int i = 0; i <= sheet.getLastRowNum(); i++) { 
            Row row = sheet.getRow(i); 
            if (row != null) { 
                Cell cell = row.getCell(0); 

                if (cell != null && cell.getStringCellValue().equalsIgnoreCase(value)) { 
                    return i; 
                } 
            } 
        } 

        return -1; 
    }
}
