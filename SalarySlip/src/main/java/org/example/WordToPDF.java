package org.example;

import com.itextpdf.text.Document;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.xwpf.usermodel.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

public class WordToPDF {

    public void convertWordToPDF(String wordFilePath, String pdfFilePath) throws Exception {
        // Load the Word document
        File wordFile = new File(wordFilePath);
        try (FileInputStream fis = new FileInputStream(wordFile);
             XWPFDocument document = new XWPFDocument(fis);
             FileOutputStream fos = new FileOutputStream(pdfFilePath)) {

            // Initialize iText PDF document
            Document pdfDoc = new Document();
            PdfWriter.getInstance(pdfDoc, fos);
            pdfDoc.open();

            // Extract and write text from paragraphs
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                for (XWPFRun run : paragraph.getRuns()) {
                    String text = run.getText(0);
                    if (text != null) {
                        pdfDoc.add(new Paragraph(text));  // Add text to the PDF
                    }
                }
            }

            // Extract and write tables
            List<XWPFTable> tables = document.getTables();
            for (XWPFTable table : tables) {
                PdfPTable pdfTable = new PdfPTable(table.getRow(0).getTableCells().size());  // Set the number of columns

                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        String cellText = cell.getText();
                        pdfTable.addCell(cellText);  // Add cell text to PDF table
                    }
                }

                pdfDoc.add(pdfTable);  // Add table to the PDF
            }

            pdfDoc.close();
            System.out.println("PDF Created: " + pdfFilePath);
        }

        // Delete the .docx file after PDF creation
        if (wordFile.delete()) {
            System.out.println("Deleted Word file: " + wordFilePath);
        } else {
            System.out.println("Failed to delete Word file: " + wordFilePath);
        }
    }
}
