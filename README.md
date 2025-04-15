You are absolutely correct! The getCellIndex() method is not a standard method of the org.apache.poi.ss.usermodel.Row interface in Apache POI.
To achieve the same functionality (getting the column index by the cell value in the header row), you need to iterate through the cells of the header row and compare their values with your HashMap key.
Here's the corrected part of the code within the while (iterator.hasNext()) loop:
                // Find the column index in Sheet1 that matches the key
                Row headerRow = sheet1.getRow(0);
                int columnIndex = -1;
                for (int j = 0; j < headerRow.getLastCellNum(); j++) {
                    Cell headerCell = headerRow.getCell(j);
                    if (headerCell != null && headerCell.getStringCellValue().equals(key)) {
                        columnIndex = j;
                        break;
                    }
                }

                if (columnIndex == -1) {
                    continue; // Key not found in Sheet1 header
                }

                for (int i = 1; i < sheet1.getLastRowNum() + 1; i++) {
                    Row dataRow = sheet1.getRow(i);
                    Cell dataCell = dataRow.getCell(columnIndex);
                    if (dataCell != null && dataCell.getStringCellValue().equals(value)) {
                        // Copy the row from Sheet1 to Sheet2
                        Row newRow = sheet2.createRow(sheet2.getLastRowNum() + 1);
                        for (int k = 0; k < dataRow.getLastCellNum(); k++) {
                            Cell oldCell = dataRow.getCell(k);
                            Cell newCell = newRow.createCell(k);
                            if (oldCell != null) {
                                switch (oldCell.getCellType()) {
                                    case STRING:
                                        newCell.setCellValue(oldCell.getStringCellValue());
                                        break;
                                    case NUMERIC:
                                        newCell.setCellValue(oldCell.getNumericCellValue());
                                        break;
                                    case BOOLEAN:
                                        newCell.setCellValue(oldCell.getBooleanCellValue());
                                        break;
                                    case FORMULA:
                                        newCell.setCellFormula(oldCell.getCellFormula());
                                        break;
                                    case BLANK:
                                        newCell.setCellType(oldCell.getCellType());
                                        break;
                                    default:
                                        // Handle other cell types if needed
                                        break;
                                }
                            }
                        }
                    }
                }

Key Changes:
 * Finding Column Index: Instead of headerRow.getCellIndex(key), we now iterate through the cells of the headerRow (the first row) using a for loop. We compare the string value of each header cell with the key from the HashMap. If a match is found, we store the column index and break out of the inner loop.
 * Handling Key Not Found: We still check if columnIndex remains -1 after iterating through the header row. If it does, it means the key from the HashMap was not found as a column header in Sheet1, and we continue to the next entry in the HashMap.
 * Copying Entire Row: When a matching row is found, we now iterate through all the cells of the dataRow in Sheet1 and copy their content and type to the newly created row in Sheet2. This ensures that all columns of the matching row are transferred.
Updated Complete Code:
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class HashMapToExcel {

    public static void main(String[] args) {
        // Create a HashMap with string keys and values
        Map<String, String> myHashMap = new HashMap<>();
        myHashMap.put("Name", "John Doe");
        myHashMap.put("City", "New York");

        // Load the Excel workbook
        File excelFile = new File("my_excel_file.xlsx");
        try (FileInputStream fis = new FileInputStream(excelFile);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // Get the Sheet1 and Sheet2 objects
            Sheet sheet1 = workbook.getSheet("Sheet1");
            Sheet sheet2 = workbook.getSheet("Sheet2");
            if (sheet2 == null) {
                sheet2 = workbook.createSheet("Sheet2");
            }

            // Iterate through the HashMap
            Iterator<Map.Entry<String, String>> iterator = myHashMap.entrySet().iterator();
            while (iterator.hasNext()) {
                Map.Entry<String, String> entry = iterator.next();
                String key = entry.getKey();
                String value = entry.getValue();

                // Find the column index in Sheet1 that matches the key
                Row headerRow = sheet1.getRow(0);
                int columnIndex = -1;
                if (headerRow != null) {
                    for (int j = 0; j < headerRow.getLastCellNum(); j++) {
                        Cell headerCell = headerRow.getCell(j);
                        if (headerCell != null && headerCell.getCellType() == CellType.STRING && headerCell.getStringCellValue().equals(key)) {
                            columnIndex = j;
                            break;
                        }
                    }
                }

                if (columnIndex == -1) {
                    System.out.println("Warning: Key '" + key + "' not found in Sheet1 header.");
                    continue; // Key not found in Sheet1 header
                }

                for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
                    Row dataRow = sheet1.getRow(i);
                    if (dataRow != null) {
                        Cell dataCell = dataRow.getCell(columnIndex);
                        if (dataCell != null) {
                            String cellValue = "";
                            if (dataCell.getCellType() == CellType.STRING) {
                                cellValue = dataCell.getStringCellValue();
                            } else if (dataCell.getCellType() == CellType.NUMERIC) {
                                cellValue = String.valueOf((int) dataCell.getNumericCellValue()); // Handle numeric values as strings for comparison
                            }

                            if (cellValue.equals(value)) {
                                // Copy the row from Sheet1 to Sheet2
                                Row newRow = sheet2.createRow(sheet2.getLastRowNum() + 1);
                                for (int k = 0; k < dataRow.getLastCellNum(); k++) {
                                    Cell oldCell = dataRow.getCell(k);
                                    Cell newCell = newRow.createCell(k);
                                    if (oldCell != null) {
                                        switch (oldCell.getCellType()) {
                                            case STRING:
                                                newCell.setCellValue(oldCell.getStringCellValue());
                                                break;
                                            case NUMERIC:
                                                newCell.setCellValue(oldCell.getNumericCellValue());
                                                break;
                                            case BOOLEAN:
                                                newCell.setCellValue(oldCell.getBooleanCellValue());
                                                break;
                                            case FORMULA:
                                                newCell.setCellFormula(oldCell.getCellFormula());
                                                break;
                                            case BLANK:
                                                newCell.setCellType(CellType.BLANK);
                                                break;
                                            default:
                                                // Handle other cell types if needed
                                                break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Save the workbook
            try (FileOutputStream fos = new FileOutputStream(excelFile)) {
                workbook.write(fos);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

Remember to replace "my_excel_file.xlsx" with the actual path to your Excel file. This corrected code should now function as intended.
