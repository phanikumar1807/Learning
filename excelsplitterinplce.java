package april22;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Iterator;
import java.util.LinkedHashMap; // To preserve column order
import java.util.Map;

public class ExcelSplitterPreserveColumns {

    public static void main(String[] args) {
        String filePath = "C:\\Users\\inahp\\eclipse-workspace\\ImageToExcel\\src\\main\\resources\\images\\sample.xlsx"; // The same file for input and output

        try {
            // Step 1: Read and process all data
            List<Map<String, String>> processedData = readAndProcessExcel(filePath);

            // Step 2: Write the processed data back to the same file/sheet
            writeToSameExcelSheetPreservingColumns(processedData, filePath);

            System.out.println("Excel file processed and saved successfully to: " + filePath);
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Error processing Excel file: " + e.getMessage());
        }
    }

    // Reads all data, identifies headers, and processes 'Issue' column
    public static List<Map<String, String>> readAndProcessExcel(String filePath) throws IOException {
        List<Map<String, String>> processedRows = new ArrayList<>();
        List<String> headers = new ArrayList<>(); // To store header names and their order

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0); // Assuming data is in the first sheet

            Iterator<Row> rowIterator = sheet.iterator();

            // Read Header Row
            if (rowIterator.hasNext()) {
                Row headerRow = rowIterator.next();
                for (Cell cell : headerRow) {
                    headers.add(getCellValueAsString(cell));
                }
            }

            // Find column indices for 'Key' and 'Issue'
            int keyColIndex = -1;
            int issueColIndex = -1;
            for (int i = 0; i < headers.size(); i++) {
                if ("Key".equalsIgnoreCase(headers.get(i))) {
                    keyColIndex = i;
                } else if ("Issue".equalsIgnoreCase(headers.get(i))) {
                    issueColIndex = i;
                }
            }

            if (keyColIndex == -1 || issueColIndex == -1) {
                throw new IllegalArgumentException("Excel file must contain 'Key' and 'Issue' headers.");
            }

            // Process Data Rows
            while (rowIterator.hasNext()) {
                Row currentRow = rowIterator.next();
                Map<String, String> originalRowData = new LinkedHashMap<>(); // Preserve order of other columns

                // Collect all cell data for the current row
                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = currentRow.getCell(i);
                    originalRowData.put(headers.get(i), getCellValueAsString(cell));
                }

                String key = originalRowData.get(headers.get(keyColIndex));
                String issuesString = originalRowData.get(headers.get(issueColIndex));

                if (key != null && !key.trim().isEmpty() && issuesString != null && !issuesString.trim().isEmpty()) {
                    String[] issues = issuesString.split(",");
                    for (String issue : issues) {
                        Map<String, String> newRow = new LinkedHashMap<>(originalRowData); // Copy all original data
                        newRow.put(headers.get(issueColIndex), issue.trim()); // Override 'Issue' column
                        processedRows.add(newRow);
                    }
                } else {
                    // If no issues to split, add the original row as is
                    processedRows.add(originalRowData);
                }
            }
        }
        return processedRows;
    }

    // Writes the processed data back to the same Excel sheet, preserving other columns
    public static void writeToSameExcelSheetPreservingColumns(List<Map<String, String>> data, String filePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);

            // Get headers from the first entry of processed data to maintain order
            List<String> headers = new ArrayList<>(data.isEmpty() ? List.of("Key", "Issue") : data.get(0).keySet());

            // Clear all data rows (but not the header)
            int lastRowNum = sheet.getLastRowNum();
            for (int i = 1; i <= lastRowNum; i++) { // Start from 1 to preserve header
                Row row = sheet.getRow(i);
                if (row != null) {
                    sheet.removeRow(row);
                }
            }
            // For older POI versions or specific clearing needs, iterating cells and setting to blank might be safer than removeRow
            // For example:
            // for (int i = 1; i <= lastRowNum; i++) {
            //     Row row = sheet.getRow(i);
            //     if (row != null) {
            //         for (int j = 0; j < row.getLastCellNum(); j++) {
            //             Cell cell = row.getCell(j);
            //             if (cell != null) {
            //                 cell.setBlank(); // Clear cell content
            //             }
            //         }
            //     }
            // }

            // Write updated data
            int rowNum = 1; // Start writing after the header row (row 0)
            for (Map<String, String> rowData : data) {
                Row newRow = sheet.createRow(rowNum++);
                for (int i = 0; i < headers.size(); i++) {
                    String header = headers.get(i);
                    Cell cell = newRow.createCell(i);
                    String value = rowData.get(header);
                    if (value != null) {
                        cell.setCellValue(value);
                    }
                }
            }

            // Write changes back to the same file
            try (FileOutputStream fos = new FileOutputStream(filePath)) {
                workbook.write(fos);
            }
        }
    }

    // Helper method to get cell value as string, handling different cell types
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return null;
        }
        CellType cellType = cell.getCellType();
        switch (cellType) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    // Return as string, handle potential decimals if keys are integers
                    return String.valueOf((int) cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return null;
        }
    }
}
