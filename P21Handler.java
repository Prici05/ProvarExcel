package com.example;

import org.apache.poi.ss.usermodel.*;

public class P21Handler {
    public static void handleP21(Workbook workbook, Workbook sourceWorkbook) {
        Sheet sheet = workbook.getSheetAt(0); // Assuming you are working with the first sheet.
        Sheet sourceSheet = sourceWorkbook.getSheet("EMAIL 1");

        int masterModulesColIndex = ExcelUtils.getColumnIndex(sheet, "Master_Modules");
        int masterElementsColIndex = ExcelUtils.getColumnIndex(sheet, "Master_Elements");
        
        String moduleValue = "P2.1.1 - Left aligned primary copy module with background colour options";
        
        int moduleRowIndex = findModuleRow(sheet, moduleValue, masterModulesColIndex);

        if (moduleRowIndex != -1 && masterElementsColIndex != -1) {
            Row moduleRow = sheet.getRow(moduleRowIndex);
            Cell headingCell = moduleRow.createCell(masterElementsColIndex);
            headingCell.setCellValue("heading");

            // Shift rows down to make room for body copy and CTA button
            sheet.shiftRows(moduleRowIndex + 1, sheet.getLastRowNum(), 2);

            Row bodyCopyRow = sheet.createRow(moduleRowIndex + 1);
            bodyCopyRow.createCell(masterElementsColIndex).setCellValue("bodycopy");

            Row ctaButtonRow = sheet.createRow(moduleRowIndex + 2);
            ctaButtonRow.createCell(masterElementsColIndex).setCellValue("CTAbutton");

            populateP21Content(sourceWorkbook, moduleValue, moduleRow, headingCell);
        }
    }

    private static void populateP21Content(Workbook sourceWorkbook, String moduleValue, Row moduleRow, Cell heading) {
        Sheet sourceSheet = sourceWorkbook.getSheet("EMAIL 1");
        
        for (int i = 5; i <= sourceSheet.getLastRowNum(); i++) { // Start from row index where data begins
            Row sourceDataRow = sourceSheet.getRow(i);
            if (sourceDataRow != null) {
                Cell moduleNameInSource = sourceDataRow.getCell(0); // Assuming MODULES is in column A

                if (moduleNameInSource != null && moduleNameInSource.getStringCellValue().equalsIgnoreCase(moduleValue)) {

                    // Fetch content from E column for heading and body copy
                    String headingContent =
                            getStringFromSource(sourceDataRow, 4); // Column E index is 4

                    heading.setStringValue(headingContent);

                    break; // Exit once we've found our match.
                }
            }
        }
    }

    private static String getStringFromSource(Row row, int columnIdx) {
        Cell cellInSourceDataColumnE = row.getCell(columnIdx);
        return cellInSourceDataColumnE != null ? cellInSourceDataColumnE.toString() : "";
    }

    private static int findModuleRow(Sheet sheet, String value, int columnIdx) {
        for (int i = 1; i <= sheet.getLastRowNum(); i++) { 
            Row rowInMasterModulesColumnA=sheet.getrow(i); 
            if(rowInMasterModulesColumnA!=null){ 
              Cell cell=rowInMasterModulesColumnA.Getcell(columnIdx); 
              if(cell!=null&&cell.Getstringcellvalue().equalsIgnoreCase(value)){ 
                 return i; 
              } 
           } 
         } 
         return -1; 
     } 
}
