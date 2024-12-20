package com.example;

import org.apache.poi.ss.usermodel.*;

public class PreHeaderHandler {
    public static void handlePreHeader(Workbook workbook, Workbook sourceWorkbook) {
        Sheet sheet = workbook.getSheetAt(0); // Assuming you are working with the first sheet.
        Sheet sourceSheet = sourceWorkbook.getSheet("EMAIL 1");

        int masterModulesColIndex = ExcelUtils.getColumnIndex(sheet, "Master_Modules");
        int masterElementsColIndex = ExcelUtils.getColumnIndex(sheet, "Master_Elements");
        int sourceModulesColumnIndex = ExcelUtils.getColumnIndex(sourceSheet, "MODULES");
        
        if (masterModulesColIndex != -1 && masterElementsColIndex != -1) {
            int preHeaderRowIndex = findRowByValue(sheet, "Pre Header", masterModulesColIndex);
            String preHeaderContent = "";

            for (int i = 5; i <= sourceSheet.getLastRowNum(); i++) { 
                Row sourceRow = sourceSheet.getRow(i);
                if (sourceRow != null) {
                    Cell moduleCell = sourceRow.getCell(sourceModulesColumnIndex); // Assuming MODULES is in column A
                    if (moduleCell != null && moduleCell.getStringCellValue().equalsIgnoreCase("Pre Header")) {
                        Cell contentCell = sourceRow.getCell(4); // Assuming E column is index 4
                        if (contentCell != null && contentCell.getCellType() == CellType.STRING) {
                            preHeaderContent = contentCell.getStringCellValue();
                            System.out.printf("PREHEADER CONTENT", preHeaderContent);
                        }
                        break;
                    }
                }
            }

            if (preHeaderRowIndex != -1) {
                Row preHeaderRow = sheet.getRow(preHeaderRowIndex);
                Cell psCell = preHeaderRow.createCell(masterElementsColIndex);
                psCell.setCellValue("ps");

                sheet.shiftRows(preHeaderRowIndex + 1, sheet.getLastRowNum(), 2);

                Row sslRow = sheet.createRow(preHeaderRowIndex + 1);
                sslRow.createCell(masterElementsColIndex).setCellValue("ssl");

                Row voRow = sheet.createRow(preHeaderRowIndex + 2);
                voRow.createCell(masterElementsColIndex).setCellValue("vo");

                populateSGENContent(sheet, preHeaderContent, masterElementsColIndex, "ssl");
                populateSGENContent(sheet, preHeaderContent, masterElementsColIndex, "vo");
            }
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
                        if ("ssl".equalsIgnoreCase(elementName)) {
                            sgEnContentCell.setCellValue(content); // Use the provided content
                        } else if ("vo".equalsIgnoreCase(elementName)) {
                            sgEnContentCell.setCellValue("View Online"); // Set value explicitly
                        }
                    }
                    break;
                }
            }
        }
    }

    private static int findRowByValue(Sheet sheet, String value, int columnIdx) {
        for (int i = 1; i <= sheet.getLastRowNum(); i++) { 
            Row rowInMasterModulesColumnA=sheet.getRow(i); 
            if(rowInMasterModulesColumnA!=null){ 
              Cell cell=rowInMasterModulesColumnA.getCell(columnIdx); 
              if(cell!=null&&cell.getStringCellValue().equalsIgnoreCase(value)){ 
                 return i; 
              } 
           } 
         } 
         return -1; 
     } 
}
