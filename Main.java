package com.example;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        String sourceFilePath = "Program Brief - Send to Receive - Asset.xlsx";
        Workbook workbook = new XSSFWorkbook();

        try (FileInputStream fis = new FileInputStream(sourceFilePath)) {
            Workbook sourceWorkbook = new XSSFWorkbook(fis);
            ExcelUtils.createHeaders(workbook, sourceWorkbook);
            
            SubjectLineHandler.handleSubjectLine(workbook, sourceWorkbook);
            PreHeaderHandler.handlePreHeader(workbook, sourceWorkbook);
            P21Handler.handleP21(workbook, sourceWorkbook);

            try (FileOutputStream fos = new FileOutputStream("ProvarExcel.xlsx")) {
                workbook.write(fos);
                System.out.println("Excel file created successfully");
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
