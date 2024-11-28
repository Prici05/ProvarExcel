package com.example;

import org.apache.poi.ss.usermodel.*;

public class P21Handler {
    public static void handleP21(Workbook workbook, Workbook sourceWorkbook) {
        Sheet sheet = workbook.getSheetAt(0); // Assuming you are working with the first sheet.
        Sheet sourceSheet = sourceWorkbook.getSheet("EMAIL 1");

        int masterModulesColIndex = ExcelUtils.getColumnIndex(sheet, "Master_Modules");
        int masterElementsColIndex = ExcelUtils.getColumnIndex(sheet, "Master_Elements");
        int sourceModulesColumnIndex = ExcelUtils.getColumnIndex(sourceSheet, "MODULES");
        
        String moduleValue = "P2.1.1 - Left aligned primary copy module with background colour options";
        
        int moduleRowIndex = findModuleByName(sheet, moduleValue);
        String pContent = "";
        String bodycopycontent = "";

        for (int i = 5; i <= sourceSheet.getLastRowNum(); i++) { 
            Row sourceRow = sourceSheet.getRow(i);
            Row bodycopyrow = sourceSheet.getRow(i+1);
            if (sourceRow != null) {
                Cell moduleCell = sourceRow.getCell(sourceModulesColumnIndex); // Assuming MODULES is in column A
                if (moduleCell != null && moduleCell.getStringCellValue().equalsIgnoreCase(moduleValue)) {
                    Cell contentCell = sourceRow.getCell(4);
                    Cell bodycopycell = bodycopyrow.getCell(4);
                     // Assuming E column is index 4
                
                    if (contentCell != null && contentCell.getCellType() == CellType.STRING) {
                        pContent = contentCell.getStringCellValue();
                        bodycopycontent = bodycopycell.getStringCellValue();
                        System.out.println("PCONTENT CONTENT " +pContent);
                        System.out.println("BODYCOPY CONTENT " +bodycopycontent);
                    }
                    break;
                }
            }
        }

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

            populateSGENContent(sheet, pContent, masterElementsColIndex, "heading");
            populateSGENContent(sheet, bodycopycontent, masterElementsColIndex, "bodycopy");

        }
    }

    private static void populateSGENContent(Sheet sheet, String content, int masterElementsColIndex, String elementName) {
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row currentRow = sheet.getRow(i);
            if (currentRow != null) {
                Cell masterElementCell = currentRow.getCell(masterElementsColIndex);
                if (masterElementCell != null && masterElementCell.getCellType() == CellType.STRING 
                    && masterElementCell.getStringCellValue().equalsIgnoreCase(elementName)) {

                    int sgEnContentColIndex = ExcelUtils.getColumnIndex(sheet, "SG-EN_Content");
                    if (sgEnContentColIndex != -1) {
                        Cell sgEnContentCell = currentRow.createCell(sgEnContentColIndex);
                        sgEnContentCell.setCellValue(content); // Use the provided content
                        
                       
                    }
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
