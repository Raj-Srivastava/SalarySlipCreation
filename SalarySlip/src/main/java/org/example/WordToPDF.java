package org.example;

import com.itextpdf.text.*;
import com.itextpdf.text.Document;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.xwpf.usermodel.*;
import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class WordToPDF {

    public void convertWordToPDF(String wordFilePath, String pdfFilePath) throws Exception {
        // Load the Word document
        File wordFile = new File(wordFilePath);
        Document pdfDoc;

        try (FileInputStream fis = new FileInputStream(wordFile);
             XWPFDocument document = new XWPFDocument(fis);
             FileOutputStream fos = new FileOutputStream(pdfFilePath)) {

            // Initialize iText PDF document
            pdfDoc = new Document();
            PdfWriter.getInstance(pdfDoc, fos);
            pdfDoc.open();

            // Extract and write text from paragraphs
            for (XWPFParagraph paragraph : document.getParagraphs()) {
                Paragraph pdfParagraph = new Paragraph();
                for (XWPFRun run : paragraph.getRuns()) {
                    String text = run.getText(0);
                    if (text != null) {
                        Font font = new Font();

                        // Apply formatting based on the Word document
                        if (run.isBold()) {
                            font.setStyle(Font.BOLD);
                        }
                        if (run.isItalic()) {
                            font.setStyle(Font.ITALIC);
                        }
                        if (run.getFontSize() != -1) {
                            font.setSize(run.getFontSize());
                        }
                        pdfParagraph.add(new Phrase(text, font));
                    }
                }
                pdfDoc.add(pdfParagraph);
            }

//            // Extract and write tables
//            List<XWPFTable> tables = document.getTables();
//            for (XWPFTable table : tables) {
//                PdfPTable pdfTable = new PdfPTable(table.getRow(0).getTableCells().size());
//                for (XWPFTableRow row : table.getRows()) {
//                    for (XWPFTableCell cell : row.getTableCells()) {
//                        String cellText = cell.getText();
//                        PdfPCell pdfCell = new PdfPCell(new Phrase(cellText));
//                        pdfCell.setPadding(5);
//                        pdfCell.setHorizontalAlignment(Element.ALIGN_CENTER);
//                        pdfTable.addCell(pdfCell);
//                    }
//                }
//                pdfDoc.add(pdfTable);
//            }
//
                pdfDoc.close(); // Close the PDF document before deleting the Word file
                System.out.println("PDF Created: " + pdfFilePath);
            }

            // Delete the .docx file after the PDF creation is successful
            if (wordFile.delete()) {
                System.out.println("Deleted Word file: " + wordFilePath);
            } else {
                System.out.println("Failed to delete Word file: " + wordFilePath);
            }
    }
}