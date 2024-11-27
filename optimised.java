package com.example;

import java.util.*;
import java.util.stream.*;
import java.io.*;
import java.util.concurrent.atomic.AtomicInteger;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ProvarExcel {

    private static final String MODULES_HEADER = "MODULES";
    private static final String MASTER_MODULES = "Master_Modules";
    private static final String MASTER_ELEMENTS = "Master_Elements";
    private static final String SG_EN_CONTENT = "SG_EN_Content";

    public static void main(String[] args) throws IOException {
        String sourceFilePath = "Program Brief - Send to Receive - Asset.xlsx";
        String outputFilePath = "ProvarExcel.xlsx";

        try (Workbook sourceWorkbook = new XSSFWorkbook(new FileInputStream(sourceFilePath));
             Workbook destWorkbook = new XSSFWorkbook()) {

            Sheet sourceSheet = sourceWorkbook.getSheet("EMAIL 1");
            Sheet destSheet = destWorkbook.createSheet("Sheet1");

            // Step 1: Create headers
            String[] headers = createHeaders(sourceSheet);
            populateHeaders(destSheet, headers);

            // Step 2: Extract module data
            int moduleColumnIndex = findColumnIndex(sourceSheet.getRow(4), MODULES_HEADER);
            List<String> moduleInputs = extractColumnData(sourceSheet, moduleColumnIndex, 5);

            // Step 3: Populate Master_Modules column
            int masterModulesColIndex = findColumnIndex(destSheet.getRow(0), MASTER_MODULES);
            populateColumn(destSheet, masterModulesColIndex, moduleInputs);

            // Step 4: Add Pre Header specific data
            addPreHeaderData(destSheet, masterModulesColIndex, headers);

            // Step 5: Populate SG_EN Content
            populateSGEnContent(destSheet, sourceSheet, masterModulesColIndex, moduleColumnIndex, headers);

            // Step 6: Write to output file
            try (FileOutputStream fos = new FileOutputStream(outputFilePath)) {
                destWorkbook.write(fos);
                System.out.println("Excel file created successfully");
            }
        }
    }

    private static String[] createHeaders(Sheet sourceSheet) {
        Row row = sourceSheet.getRow(3);
        String prefix = row.getCell(4).getStringCellValue().split(" ")[0];
        return Stream.concat(
                Arrays.stream(new String[]{MASTER_MODULES, "Module_Background_Color", MASTER_ELEMENTS}),
                Arrays.stream(new String[]{prefix + "_Content", prefix + "_Link"})
        ).toArray(String[]::new);
    }

    private static void populateHeaders(Sheet sheet, String[] headers) {
        Row headerRow = sheet.createRow(0);
        AtomicInteger index = new AtomicInteger(0);
        Arrays.stream(headers).forEach(header -> {
            Cell cell = headerRow.createCell(index.getAndIncrement());
            cell.setCellValue(header);
        });
    }

    private static int findColumnIndex(Row row, String columnName) {
        for (Cell cell : row) {
            if (cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                return cell.getColumnIndex();
            }
        }
        return -1;
    }

    private static List<String> extractColumnData(Sheet sheet, int columnIndex, int startRow) {
        List<String> data = new ArrayList<>();
        for (int i = startRow; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    data.add(cell.getStringCellValue());
                }
            }
        }
        return data;
    }

    private static void populateColumn(Sheet sheet, int columnIndex, List<String> data) {
        for (int i = 0; i < data.size(); i++) {
            Row row = sheet.createRow(i + 1);
            Cell cell = row.createCell(columnIndex);
            cell.setCellValue(data.get(i));
        }
    }

    private static void addPreHeaderData(Sheet sheet, int masterModulesColIndex, String[] headers) {
        int masterElementsColIndex = findColumnIndex(sheet.getRow(0), MASTER_ELEMENTS);
        int preHeaderRowIndex = findRowByValue(sheet, masterModulesColIndex, "Pre Header");

        if (preHeaderRowIndex != -1 && masterElementsColIndex != -1) {
            Row preHeaderRow = sheet.getRow(preHeaderRowIndex);
            preHeaderRow.createCell(masterElementsColIndex).setCellValue("ps");

            sheet.shiftRows(preHeaderRowIndex + 1, sheet.getLastRowNum(), 2);

            sheet.createRow(preHeaderRowIndex + 1).createCell(masterElementsColIndex).setCellValue("ssl");
            sheet.createRow(preHeaderRowIndex + 2).createCell(masterElementsColIndex).setCellValue("vo");
        }
    }

    private static int findRowByValue(Sheet sheet, int columnIndex, String value) {
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell cell = row.getCell(columnIndex);
                if (cell != null && cell.getCellType() == CellType.STRING && cell.getStringCellValue().equalsIgnoreCase(value)) {
                    return i;
                }
            }
        }
        return -1;
    }

    private static void populateSGEnContent(Sheet destSheet, Sheet sourceSheet, int masterModulesColIndex, int moduleColumnIndex, String[] headers) {
        int sgEnContentColIndex = findColumnIndex(destSheet.getRow(0), SG_EN_CONTENT);

        for (int i = 1; i <= destSheet.getLastRowNum(); i++) {
            Row destRow = destSheet.getRow(i);
            if (destRow != null) {
                Cell masterModuleCell = destRow.getCell(masterModulesColIndex);
                if (masterModuleCell != null && masterModuleCell.getCellType() == CellType.STRING) {
                    String moduleValue = masterModuleCell.getStringCellValue();
                    String content = findContentForModule(sourceSheet, moduleColumnIndex, moduleValue);
                    if (content != null && sgEnContentColIndex != -1) {
                        destRow.createCell(sgEnContentColIndex).setCellValue(content);
                    }
                }
            }
        }
    }

    private static String findContentForModule(Sheet sheet, int moduleColumnIndex, String moduleValue) {
        for (int i = 5; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                Cell moduleCell = row.getCell(moduleColumnIndex);
                if (moduleCell != null && moduleCell.getCellType() == CellType.STRING &&
                        moduleCell.getStringCellValue().equalsIgnoreCase(moduleValue)) {
                    Cell contentCell = row.getCell(4); // Assuming content is in column E (index 4)
                    return contentCell != null && contentCell.getCellType() == CellType.STRING
                            ? contentCell.getStringCellValue() : null;
                }
            }
        }
        return null;
    }
}