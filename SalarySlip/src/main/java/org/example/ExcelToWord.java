package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;

public class ExcelToWord {
    public void convertExcelToWord(String excelFilePath, String wordFilePath) throws Exception {
        // Load the Excel file
        FileInputStream excelFile = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet sheet = workbook.getSheetAt(0);

        // Iterate over the rows in the Excel sheet
        Iterator<Row> iterator = sheet.iterator();
        // Skip the first row (header)
        if (iterator.hasNext()) {
            iterator.next();
        }

        // Process each row
        while (iterator.hasNext()) {
            Row row = iterator.next();

            // Get combined name from the first cell
            String fullName = row.getCell(0) != null ? row.getCell(0).toString() : "";
            String[] nameParts = fullName.split(" ", 2); // Split only on the first space
            String firstName = nameParts.length > 0 ? nameParts[0] : "";
            String lastName = nameParts.length > 1 ? nameParts[1] : "";

            // Get additional values from the next cells
            String designation = row.getCell(1) != null ? row.getCell(1).toString() : "";
            String pan = row.getCell(7) != null ? row.getCell(7).toString() : "";

            // Load a Word document template
            FileInputStream templateFile = new FileInputStream(wordFilePath);
            XWPFDocument doc = new XWPFDocument(templateFile);

            // Replace placeholders in the Word document with actual values
            for (XWPFParagraph paragraph : doc.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    String text = run.getText(0);
                    if (text != null && !text.isEmpty()) {
                        // Replace placeholders with actual values
                        text = text.replace("<First>", firstName)
                                .replace("<Last>", lastName)
                                .replace("<Designation>", designation)
                                .replace("<PAN>", pan);
                        run.setText(text, 0); // Update the text in the run
                    }
                }
            }

            // Save the updated Word document with a unique name for each row
            String outputFileName = firstName + "_" + lastName + "_output.docx"; // Unique name
            String fullPath = "src/resources/" + outputFileName; // Ensure this path exists
            try (FileOutputStream out = new FileOutputStream(fullPath)) {
                doc.write(out);
            }
            templateFile.close(); // Close the template file after use

            // Log the output file creation
            System.out.println("Word Created at: " + fullPath);
        }

        // Close the workbook and Excel file
        workbook.close();
        excelFile.close();
    }
}
