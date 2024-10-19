package org.example;

import org.apache.pdfbox.pdmodel.PDDocument;

import java.io.File;

public class WordToPDF {
    public void convertWordToPDF(String wordFilePath, String pdfFilePath) throws Exception {
        // Load the Word document
        File wordFile = new File(wordFilePath);

        // Convert to PDF using PDFBox or similar
        PDDocument pdfDocument = PDDocument.load(wordFile);

        // Save as PDF
        pdfDocument.save(pdfFilePath);
        pdfDocument.close();
    }
}
