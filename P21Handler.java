package com.example;

import org.apache.poi.ss.usermodel.*;

public class P21Handler {
    public static void handleP21(Workbook workbook, Workbook sourceWorkbook) {
        Sheet sheet = workbook.getSheetAt(0); // Assuming you are working with the first sheet.
        Sheet sourceSheet = sourceWorkbook.getSheet("EMAIL 1");

        int masterModulesColIndex = ExcelUtils.getColumnIndex(sheet, "Master_Modules");
        int masterElementsColIndex = ExcelUtils.getColumnIndex(sheet, "Master_Elements");
        
        
        String moduleValue ="P2.1.1 - Left aligned primary copy module with background colour options";
        
        int moduleRowIndex= findModuleByName(sheet,moduleValue);
     

       if(moduleRowIndex!=-1 && masterElementsColIndex!=-1){
           Row modulrow=sheet.getRow(moduleRowIndex);
           Cell headingcell=modulrow.createCell(masterElementsColIndex);
           headingcell.setCellValue("heading");

           // Shift rows below the P2.1.1 row down by two positions to make room for body copy and CTA button.
           sheet.shiftRows(moduleRowIndex+1,sheet.getLastRowNum(),2);

           Row bodycopyrow=sheet.createRow(moduleRowIndex+1);
           bodycopyrow.createCell(masterElementsColIndex).setCellValue("bodycopy");

           Row ctbuttonrow=sheet.createRow(moduleRowIndex+2);
           ctbuttonrow.createCell(masterElementsColIndex).setCellValue("CTAbutton");

           populateP21Content(sourceWorkbook,moduleValue,modulrow, sheet);

       }

   }

   private static void populateP21Content(Workbook sourceWorkbook,String moduleValue,Row modulrow, Sheet sheet){
    
       Sheet sourcesheet=sourceWorkbook.getSheet("EMAIL 1");
       int sourceModulesColumnIndex = ExcelUtils.getColumnIndex(sourcesheet, "MODULES");
       int sgEnContentColIndex = ExcelUtils.getColumnIndex(sheet, "SG-EN_Content");
       
       for(int i=1;i<=sourcesheet.getLastRowNum();i++){ 
          Row sourcedatarow=sourcesheet.getRow(i); 
          if(sourcedatarow!=null){ 
             Cell modulenameinsourcedata=sourcedatarow.getCell(sourceModulesColumnIndex); 
             System.out.println("**********" +modulenameinsourcedata);

             if(modulenameinsourcedata!=null&&modulenameinsourcedata.getStringCellValue().equalsIgnoreCase(moduleValue)){ 

                 // Fetch content from E column for heading and body copy.
                 String headingcontent=sourcedatarow.getCell(4)!=null?sourcedatarow.getCell(4).getStringCellValue():""; 
                 System.out.println("HEADING CONTENT " +headingcontent);
                 modulrow.createCell(sgEnContentColIndex).setCellValue(headingcontent); 

                 break; 

             } 

          } 

       } 

   }

   private static int findModuleByName(Sheet sheet,String value){
       for(int i=0;i<=sheet.getLastRowNum();i++){ 
          Row row=sheet.getRow(i); 
          if(row!=null){ 
             Cell cell=row.getCell(0); 

             if(cell!=null&&cell.getStringCellValue().equalsIgnoreCase(value)){ 

                 return i; 

             } 

          } 

       } 

       return-1; 

   }

}
