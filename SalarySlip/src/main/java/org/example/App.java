package org.example;

public class App 
{
    public static void main(String[] args) {
        try {
            // Step 1: Convert Excel to Word document
            ExcelToWord excelToWord = new ExcelToWord();
            excelToWord.convertExcelToWord("src/resources/SalarySheetTemplateMM.xlsm", "src/resources/Salary Slip Template .docx");
            System.out.println("Excel to Word conversion completed.");

        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("An error occurred during the process.");
        }
    }
}
