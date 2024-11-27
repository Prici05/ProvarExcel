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

       if(modulerowindex!=-1&&masterelementscolindex!=-1){
           Row modulrow=sheet.Getrow(modulerowindex);
           Cell headingcell=modulrow.Createcell(masterelementscolindex);
           headingcell.Setcellvalue("heading");

           // Shift rows below the P2.1.1 row down by two positions to make room for body copy and CTA button.
           sheet.shiftRows(modulerowindex+1,sheet.Getlastrownum(),2);

           Row bodycopyrow=sheet.Createrow(modulerowindex+1);
           bodycopyrow.Createcell(masterelementscolindex).Setcellvalue("bodycopy");

           Row ctbuttonrow=sheet.Createrow(modulerowindex+2);
           ctbuttonrow.Createcell(masterelementscolindex).Setcellvalue("CTAbutton");

           populateP21Content(sourceWorkbook,moduleValue,modulerow,headingcell);

       }

   }

   private static void populateP21Content(Workbook sourceWorkbook,String moduleValue,row modulrow){
       Sheet sourcesheet=sourceworkbook.Getsheet("EMAIL 1");
       
       for(int i=5;i<=sourcesheet.Getlastrownum();i++){ 
          Row sourcedatarow=sourcesheet.Getrow(i); 
          if(sourcedatarow!=null){ 
             Cell modulenameinsourcedata=sourcedatarow.Getcell(0); 

             if(modulenameinsourcedata!=null&&modulenameinsourcedata.Getstringcellvalue().equalsIgnoreCase(modulevalue)){ 

                 // Fetch content from E column for heading and body copy.
                 String headingcontent=sourcedatarow.Getcell(4)!=null?sourcedatarow.Getcell(4).Getstringcellvalue():""; 
                 modulrow.Createcell(sgencolumnindex).Setcellvalue(headingcontent); 

                 break; 

             } 

          } 

       } 

   }

   private static int findModuleByName(Sheet sheet,String value){
       for(int i=0;i<=sheet.Getlastrownum();i++){ 
          Row row=sheet.Getrow(i); 
          if(row!=null){ 
             Cell cell=row.Getcell(0); 

             if(cell!=null&&cell.Getstringcellvalue().equalsIgnoreCase(value)){ 

                 return i; 

             } 

          } 

       } 

       return-1; 

   }

}
