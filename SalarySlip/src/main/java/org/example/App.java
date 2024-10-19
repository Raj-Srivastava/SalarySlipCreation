package org.example;

public class App 
{
    public static void main(String[] args) {
        try {
            // Step 1: Convert Excel to Word document
            ExcelToWord excelToWord = new ExcelToWord();
            excelToWord.convertExcelToWord("src/resources/SalarySheetTemplateMM.xlsm", "src/resources/Salary Slip Template .docx");
            System.out.println("Excel to Word conversion completed.");

//            // Step 2: Convert Word to PDF document
//            WordToPDF wordToPDF = new WordToPDF();
//            wordToPDF.convertWordToPDF("src/resources/Salary Slip Template .docx", "src/resources/output.pdf");
//            System.out.println("Word to PDF conversion completed.");

//            // Step 3: Send the PDF via email
//            SendEmail sendEmail = new SendEmail();
//            sendEmail.sendEmailWithAttachment("rajsrivastava@gmail.com", "output.pdf");
//            System.out.println("Email with PDF attachment sent successfully.");

        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("An error occurred during the process.");
        }
    }
}
