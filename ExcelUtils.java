package com.example;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Stream;

public class ExcelUtils {

    public static void createHeaders(Workbook workbook, Workbook sourceWorkbook) {
        Sheet sheet = workbook.createSheet("Sheet1");
        
        // Static headers
        String[] headers = {"Master_Modules", "Module_Background_Color", "Master_Elements"};
        
        // Create the first row for headers
        Row headerRow = sheet.createRow(0);
        int index = 0;

        // Add static headers
        for (String header : headers) {
            Cell cell = headerRow.createCell(index++);
            cell.setCellValue(header);
        }

        // Dynamically create additional headers based on source data
        try {
            Sheet sourceSheet = sourceWorkbook.getSheet("EMAIL 1");
            Row row4 = sourceSheet.getRow(3); // Assuming dynamic values are in row 4
            Cell cellE4 = row4.getCell(4); // Assuming dynamic value is in column E (index 4)
            String cellValue = cellE4.getStringCellValue();
            String prefix = cellValue.split(" ")[0]; // Get prefix from the cell value

            String[] dynamicHeaders = {
                prefix + "_Content",
                prefix + "_Link"
            };

            // Combine static and dynamic headers
            String[] finalHeaders = Stream.concat(Arrays.stream(headers), Arrays.stream(dynamicHeaders))
                                          .toArray(String[]::new);

            // Add combined headers to the sheet
            index = 0;
            for (String finalHeader : finalHeaders) {
                Cell cell = headerRow.createCell(index++);
                cell.setCellValue(finalHeader);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static int getColumnIndex(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0);
        for (int col = 0; col < headerRow.getLastCellNum(); col++) {
            Cell cell = headerRow.getCell(col);
            if (cell != null && cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                return col;
            }
        }
        return -1; // Column not found
    }
}
