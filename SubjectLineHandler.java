
package com.example;

import org.apache.poi.ss.usermodel.*;

public class SubjectLineHandler {
    public static void handleSubjectLine(Workbook workbook, Workbook sourceWorkbook) {
        Sheet sheet = workbook.getSheetAt(0); // Assuming you are working with the first sheet.
        Sheet sourceSheet = sourceWorkbook.getSheet("EMAIL 1");
        
        int masterModulesColIndex = ExcelUtils.getColumnIndex(sheet, "Master_Modules");
        int sgEnContentColIndex = ExcelUtils.getColumnIndex(sheet, "SG-EN_Content");
        int sourceModulesColumnIndex = ExcelUtils.getColumnIndex(sourceSheet, "MODULES");
        
        
        if (masterModulesColIndex != -1 && sgEnContentColIndex != -1) {
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row currentRow = sheet.getRow(i);
                if (currentRow != null) {
                    Cell masterModuleCell = currentRow.getCell(masterModulesColIndex);
                    if (masterModuleCell != null && masterModuleCell.getCellType() == CellType.STRING 
                        && masterModuleCell.getStringCellValue().equalsIgnoreCase("Subject Line")) {

                        for (int j = 5; j <= sourceSheet.getLastRowNum(); j++) {
                            Row sourceRow = sourceSheet.getRow(j);
                            if (sourceRow != null) {
                                Cell moduleCell = sourceRow.getCell(sourceModulesColumnIndex); // Assuming MODULES is in column A (index 0)
                                if (moduleCell != null && moduleCell.getCellType() == CellType.STRING 
                                    && moduleCell.getStringCellValue().equalsIgnoreCase("Subject Line")) {

                                    Cell contentCell = sourceRow.getCell(4); // Assuming E column is index 4
                                    if (contentCell != null && contentCell.getCellType() == CellType.STRING) {
                                        Cell sgEnContentCell = currentRow.createCell(sgEnContentColIndex);
                                        sgEnContentCell.setCellValue(contentCell.getStringCellValue());
                                    }
                                    break;
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
