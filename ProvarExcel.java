package com.example;
import java.util.Arrays;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.stream.Stream;
import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
public class ProvarExcel {
   public static void main(String[] args) throws FileNotFoundException, IOException {
       String sourceFilePath = "Program Brief - Send to Receive - Asset.xlsx"; // Path to source file
       // 1. create a new workbook
       Workbook workbook = new XSSFWorkbook();
       // 2. Create a new sheet
       Sheet sheet = workbook.createSheet("Sheet1");
       Row row = sheet.createRow(0);
       String[] headers = {"Master_Modules", "Module_Background_Color", "Master_Elements"};
       try (FileInputStream fis = new FileInputStream(sourceFilePath)) {
           Workbook sourceWorkbook = new XSSFWorkbook(fis);
           {
               Sheet sourceSheet = sourceWorkbook.getSheet("EMAIL 1");
               String[] finalHeaders = null;
               for (int i = 0; i < sourceWorkbook.getNumberOfSheets(); i++) {
                   Row row4 = sourceSheet.getRow(3);
                   Cell cellE4 = row4.getCell(4);
                   String cellvalue = cellE4.getStringCellValue();
                   System.out.println(cellvalue);
                   String prefix = cellvalue.split(" ")[0];
                   System.out.println(prefix);
                   String[] dynamicheaders = {new StringBuilder().append(prefix).append("_Content").toString(),
                   new StringBuilder().append(prefix).append("_Link").toString()};
                   finalHeaders = Stream.concat(Arrays.stream(headers), Arrays.stream(dynamicheaders))
                           .toArray(String[]::new);
                   System.out.println(finalHeaders);
               }
               Row headingrow = sourceSheet.getRow(4);
               int reqcolindex = -1;
               for (Cell cell : headingrow) {
                   if (cell.getStringCellValue().equalsIgnoreCase("MODULES")) {
                       reqcolindex = cell.getColumnIndex();
                       break;
                   }
               }
               if (reqcolindex == -1) {
                   System.out.println("There is no column called Modules in the input file excel");
               }
               List<String> moduleinputs = new ArrayList<>();
               for (int j = 5; j <= sourceSheet.getLastRowNum(); j++) {
                   Row destRow = sourceSheet.getRow(j);
                   if (destRow != null) {
                       Cell destcell = destRow.getCell(reqcolindex);
                       if (destcell != null && destcell.getCellType() == CellType.STRING) {
                           moduleinputs.add(destcell.getStringCellValue());
                       }
                   }
               }
               AtomicInteger index = new AtomicInteger(0);
               Arrays.stream(finalHeaders).forEachOrdered(header ->
               {
                   Cell cell = row.createCell(index.getAndIncrement());
                   cell.setCellValue(header);
               });
               // Add module data to "Master_Modules" column
               int masterModulesColIndex = -1;
               for (Cell cell : row) {
                   if (cell.getStringCellValue().equalsIgnoreCase("Master_Modules")) {
                       masterModulesColIndex = cell.getColumnIndex();
                       break;
                   }
               }
               if (masterModulesColIndex != -1) {
                   for (int i = 0; i < moduleinputs.size(); i++) {
                       Row newRow = sheet.getRow(i + 1);
                       if (newRow == null) {
                           newRow = sheet.createRow(i + 1);
                       }
                       Cell newCell = newRow.createCell(masterModulesColIndex);
                       newCell.setCellValue(moduleinputs.get(i));
                   }
               } else {
                   System.out.println("Master_Modules column not found in ProvarExcel.xlsx");
               }
               // Find the row with "Pre Header" and add "ps", "ssl", and "vo" under "Master_Elements"
               int preHeaderRowIndex = -1;
               for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                   Row currentRow = sheet.getRow(i);
                   if (currentRow != null) {
                       for (Cell cell : currentRow) {
                           if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().equalsIgnoreCase("Pre Header")) {
                               preHeaderRowIndex = currentRow.getRowNum();
                               break;
                           }
                       }
                   }
                   if (preHeaderRowIndex != -1) {
                       break;
                   }
               }
               int masterElementsColIndex = -1;
               for (Cell cell : row) {
                   if (cell.getStringCellValue().equalsIgnoreCase("Master_Elements")) {
                       masterElementsColIndex = cell.getColumnIndex();
                       break;
                   }
               }
               if (preHeaderRowIndex != -1 && masterElementsColIndex != -1) {
                   // Add "ps" in the same row as "Pre Header"
                   Row preHeaderRow = sheet.getRow(preHeaderRowIndex);
                   Cell psCell = preHeaderRow.createCell(masterElementsColIndex);
                   psCell.setCellValue("ps");
                   // Add "ssl" and "vo" in the rows below
                   Row sslRow = sheet.createRow(preHeaderRowIndex + 1);
                   Cell sslCell = sslRow.createCell(masterElementsColIndex);
                   sslCell.setCellValue("ssl");
                   Row voRow = sheet.createRow(preHeaderRowIndex + 2);
                   Cell voCell = voRow.createCell(masterElementsColIndex);
                   voCell.setCellValue("vo");
               } else {
                   System.out.println("Pre Header row or Master_Elements column not found.");
               }
           }
       }
       try (FileOutputStream fos = new FileOutputStream("ProvarExcel.xlsx")) {
           workbook.write(fos);
           System.out.println("Excel file created successfully with Master_Modules and Pre Header data");
       } catch (IOException e) {
           e.printStackTrace();
       }
   }
}
