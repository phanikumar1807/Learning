package april22;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExcelSplitterInPlace {

    public static void main(String[] args) {
        String filePath = ""; // The same file for input and output

        try {
            // Step 1: Read and process the data
            List<RowData> processedData = readAndProcessExcel(filePath);

            // Step 2: Write the processed data back to the same file/sheet
            writeToSameExcelSheet(processedData, filePath);

            System.out.println("Excel file processed and saved successfully to: " + filePath);
        } catch (IOException e) {
            e.printStackTrace();
            System.err.println("Error processing Excel file: " + e.getMessage());
        }
    }

    public static List<RowData> readAndProcessExcel(String filePath) throws IOException {
        List<RowData> processedRows = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0); // Assuming data is in the first sheet

            boolean isFirstRow = true;
            for (Row row : sheet) {
                if (isFirstRow) {
                    isFirstRow = false;
                    continue; // Skip header row during processing
                }

                Cell keyCell = row.getCell(0); // Assuming Key is in the first column (index 0)
                Cell issueCell = row.getCell(1); // Assuming Issue is in the second column (index 1)

                if (keyCell == null || issueCell == null) {
                    continue; // Skip rows with missing key or issue data
                }

                String key = getCellValueAsString(keyCell);
                String issuesString = getCellValueAsString(issueCell);

                if (key != null && !key.trim().isEmpty() && issuesString != null && !issuesString.trim().isEmpty()) {
                    String[] issues = issuesString.split(",");
                    for (String issue : issues) {
                        processedRows.add(new RowData(key, issue.trim())); // Trim whitespace
                    }
                }
            }
        } // Workbook is closed here, but we need to re-open it for writing
        return processedRows;
    }

    public static void writeToSameExcelSheet(List<RowData> data, String filePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) { // Re-open the workbook for writing
            Sheet sheet = workbook.getSheetAt(0);

            // Clear existing rows in the sheet
            int lastRowNum = sheet.getLastRowNum();
            for (int i = 0; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    sheet.removeRow(row);
                }
            }

            // Create header row
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Key");
            headerRow.createCell(1).setCellValue("Issue");

            int rowNum = 1; // Start writing from the second row after the header
            for (RowData rowData : data) {
                Row row = sheet.createRow(rowNum++);
                row.createCell(0).setCellValue(rowData.getKey());
                row.createCell(1).setCellValue(rowData.getIssue());
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

    // Simple data class to hold processed row data
    private static class RowData {
        private String key;
        private String issue;

        public RowData(String key, String issue) {
            this.key = key;
            this.issue = issue;
        }

        public String getKey() {
            return key;
        }

        public String getIssue() {
            return issue;
        }
    }
}
