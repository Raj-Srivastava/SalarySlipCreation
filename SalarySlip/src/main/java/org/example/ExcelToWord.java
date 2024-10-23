package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class ExcelToWord {
    public void convertExcelToWord(String excelFilePath, String wordFilePath) throws Exception {
        // Load the Excel file
        FileInputStream excelFile = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet sheet = workbook.getSheetAt(0);

        // Defining Arrays of Placeholder and column indexes
        String[] placeholders = {"<First>","<Last>","<Designation>","<PAN>","<Account>"};
        int[] columnIndexes = {1,1,2,8,9};
        // Iterate over the rows in the Excel sheet
        Iterator<Row> iterator = sheet.iterator();
        while (iterator.hasNext()) {
            Row row = iterator.next();

            // Skip processing the row if it's the first (header) row
            if (row.getRowNum() == 0) {
                continue;
            }

            // Array to store values dynamically based on the defined mapping
            String[] values = new String[placeholders.length];

            // Read the cells based on the defined mapping
            for (int i = 0; i < placeholders.length; i++) {
                // Only process if the index exists in the row
                if (i < columnIndexes.length) {
                    Cell cell = row.getCell(columnIndexes[i]);
                    if (cell != null) {
                        values[i] = cell.toString(); // Handle different cell types
                    } else {
                        values[i] = ""; // Handle case where the cell is null
                    }
                } else {
                    values[i] = ""; // Handle case where column index is out of bounds
                }
            }

            // Split the full name into first and last names if the first value exists
            String fullName = values[0]; // Full name from the first index
            String firstName = "";
            String lastName = "";

            if (!fullName.isEmpty()) {
                String[] nameParts = fullName.split(" ", 2); // Split on the first space
                firstName = nameParts[0]; // Get first name
                lastName = nameParts.length > 1 ? nameParts[1] : ""; // Get last name if available
            }

            // Update the values array to include first and last names
            values[0] = firstName; // Update first index with first name
            if (placeholders.length > 1) {
                values[1] = lastName; // Update second index with last name
            }

            // Load the Word document template for each row to ensure a fresh template
            try (FileInputStream templateFile = new FileInputStream(wordFilePath)) {
                XWPFDocument doc = new XWPFDocument(templateFile);

                // **Replace Placeholders in Paragraphs (outside tables)**
                for (XWPFParagraph paragraph : doc.getParagraphs()) {
                    for (XWPFRun run : paragraph.getRuns()) {
                        String text = run.getText(0);
                        if (text != null && !text.isEmpty()) {
                            // Replace each placeholder dynamically based on the values array
                            for (int i = 0; i < placeholders.length; i++) {
                                text = text.replace(placeholders[i], values[i]);
                            }
                            run.setText(text, 0); // Update the text in the run
                        }
                    }
                }

                // **Replace Placeholders in Tables (if placeholders are inside tables)**
                for (XWPFTable table : doc.getTables()) {
                    for (XWPFTableRow tableRow : table.getRows()) {
                        for (XWPFTableCell cellTable : tableRow.getTableCells()) {
                            for (XWPFParagraph paragraphTable : cellTable.getParagraphs()) {
                                for (XWPFRun run : paragraphTable.getRuns()) {
                                    String text = run.getText(0);
                                    if (text != null && !text.isEmpty()) {
                                        // Replace each placeholder dynamically based on the values array
                                        for (int i = 0; i < placeholders.length; i++) {
                                            text = text.replace(placeholders[i], values[i]);
                                        }
                                        run.setText(text, 0); // Update the text in the run
                                    }
                                }
                            }
                        }
                    }
                }

                // Save the updated Word document with a unique name for each row
                String outputFileName = firstName + " " + lastName + " Salary Slip.docx"; // Unique name based on first and last name
                String fullPath = "src/resources/" + outputFileName; // Ensure this path exists
                try (FileOutputStream out = new FileOutputStream(fullPath)) {
                    doc.write(out);
                }
                System.out.println("Word Created at: " + fullPath);
            } catch (Exception e) {
                e.printStackTrace(); // Handle exceptions for file operations
            }
        }
        workbook.close();
        excelFile.close();
    }
}
