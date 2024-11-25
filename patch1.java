 String preHeaderContent = "";
           for (int i = 5; i <= sourceSheet.getLastRowNum(); i++) {
               Row sourceRow = sourceSheet.getRow(i);
               if (sourceRow != null) {
                   Cell moduleCell = sourceRow.getCell(reqcolindex);
                   System.out.println("MODULE CELL" +moduleCell);
                   if (moduleCell != null && moduleCell.getCellType() == CellType.STRING &&
                           moduleCell.getStringCellValue().equalsIgnoreCase("Pre Header")) {
                       // Get the content from column E (index 4)
                       Cell contentCell = sourceRow.getCell(4);
                       System.out.println("PREHEADER CONTENT" +contentCell); // Assuming column E contains the content
                       if (contentCell != null && contentCell.getCellType() == CellType.STRING) {
                           preHeaderContent = contentCell.getStringCellValue();
                       }
                       break;
                   }
               }
           }

              // Now populate SG_EN Content for "ssl" with the fetched Pre Header content
              if (preHeaderContent != null && !preHeaderContent.isEmpty()) {
                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row currentRow = sheet.getRow(i);
                    if (currentRow != null) {
                        Cell masterElementCell = currentRow.getCell(masterElementsColIndex);
                        if (masterElementCell != null && masterElementCell.getCellType() == CellType.STRING &&
                        masterElementCell.getStringCellValue().equalsIgnoreCase("ssl")) {
                            // Set SG_EN Content for "ssl"
                            
                            for (Cell cell : row) {
                                if (cell.getStringCellValue().equalsIgnoreCase("SG_EN Content")) {
                                    sgEnContentColIndex = cell.getColumnIndex();
                                    break;
                                }
                            }
                            if (sgEnContentColIndex != -1) {
                                Cell sgEnContentCell = currentRow.createCell(sgEnContentColIndex);
                                sgEnContentCell.setCellValue(preHeaderContent);
                            }
                        }
                    }
                }
            }
