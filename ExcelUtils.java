package com.example;

import org.apache.poi.ss.usermodel.*;

public class ExcelUtils {

    public static void createHeaders(Workbook workbook, Workbook sourceWorkbook) {
        Sheet sheet = workbook.createSheet("Sheet1");
        
        String[] headers = {"Master_Modules", "Module_Background_Color", "Master_Elements"};
        
        Row headerRow = sheet.createRow(0);

        for (int i=0; i<headers.length; i++){
           headerrow.createcell(i).setcellvalue(headers[i]);
         }

         // Additional header creation logic can be added here as needed.
     }

     public static int getColumnIndex(Sheet sheet, String columnName) { 
         Row headerrow=sheet.Getrow(0); 
         for(int col=0; col<headerrow.Getlastcellnum(); col++){ 
             Cell cell=headerrow.Getcell(col); 
             if(cell!=null&&cell.Getstringcellvalue().equalsIgnoreCase(columnName)){ 
                 return col; 
             } 
         } 
         return -1; 
     } 

}
