package com.example;

import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

public class ModuleHandler {
    public static void handleModules(Workbook workbook, Workbook sourceWorkbook) {
        Sheet sheet = workbook.getSheetAt(0); // Assuming you are working with the first sheet.
        Sheet sourceSheet = sourceWorkbook.getSheet("EMAIL 1");

        int masterModulesColIndex = ExcelUtils.getColumnIndex(sheet, "Master_Modules");
        
        if (masterModulesColIndex != -1) {
            List<String> moduleInputs = new ArrayList<>();
            int reqColIndex = findColumnIndex(sourceSheet, "MODULES");

            // Fetch values under MODULES column
            for (int j = 5; j <= sourceSheet.getLastRowNum(); j++) { // Start from row 6
                Row destRow = sourceSheet.getRow(j);
                if (destRow != null) {
                    Cell destCell = destRow.getCell(reqColIndex);
                    if (destCell != null && destCell.getCellType() == CellType.STRING) {
                        moduleInputs.add(destCell.getStringCellValue());
                    }
                }
            }

            // Add module data to "Master_Modules" column in ProvarExcel
            for (int i = 0; i < moduleInputs.size(); i++) {
                Row newRow = sheet.createRow(i + 1); // Start adding from row 2
                Cell newCell = newRow.createCell(masterModulesColIndex);
                newCell.setCellValue(moduleInputs.get(i));
            }
        }
    }

    private static int findColumnIndex(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(4); // Assuming headers are in row 5
        for (int col = 0; col < headerRow.getLastCellNum(); col++) {
            Cell cell = headerRow.getCell(col);
            if (cell != null && cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                return col;
            }
        }
        return -1; // Column not found
    }
}
