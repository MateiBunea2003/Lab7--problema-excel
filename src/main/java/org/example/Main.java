package org.example;/*package org.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


// Java Program to Illustrate Reading
// Data to Excel File Using Apache POI

// Import statements


// Main class
public class Main {

    // Main driver method
    public static void main(String[] args) {

        // Try block to check for exceptions
        try {

            // Reading file from local directory
            FileInputStream file = new FileInputStream(
                    new File("Lab7.xlsx"));

            // Create Workbook instance holding reference to
            // .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);

            // Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            // Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();

            // Till there is an element condition holds true
            while (rowIterator.hasNext()) {

                Row row = rowIterator.next();

                // For each row, iterate through all the
                // columns
                Iterator<Cell> cellIterator
                        = row.cellIterator();

                while (cellIterator.hasNext()) {

                    Cell cell = cellIterator.next();

                    // Checking the cell type and format
                    // accordingly
                    switch (cell.getCellType()) {

                        // Case 1
                        case NUMERIC:
                            System.out.print(
                                    cell.getNumericCellValue()
                                            + "\t");
                            break;

                        // Case 2
                        case STRING:
                            System.out.print(
                                    cell.getStringCellValue()
                                            + "\t");
                        default:
                            break;
                        case FORMULA:
                            System.out.print(
                                    cell.getNumericCellValue()
                                            + "\t");
                    }
                }

                System.out.println("");
            }

            // Closing file output streams
            file.close();
        }

        // Catch block to handle exceptions
        catch (Exception e) {

            // Display the exception along with line number
            // using printStackTrace() method
            e.printStackTrace();
        }
    }
}*/

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

    public static void main(String[] args) {

        try {
            FileInputStream file = new FileInputStream("Lab7.xlsx");
            XSSFWorkbook inputWorkbook = new XSSFWorkbook(file);
            XSSFSheet inputSheet = inputWorkbook.getSheetAt(0);

            XSSFWorkbook outputWorkbook = new XSSFWorkbook();
            XSSFSheet outputSheet = outputWorkbook.createSheet("Lab7-out");

            FormulaEvaluator formulaEvaluator = inputWorkbook.getCreationHelper().createFormulaEvaluator();

            int rowNum = 0;
            for (Row row : inputSheet) {
                Row outputRow = outputSheet.createRow(rowNum++);
                Cell cell1 = row.getCell(0);
                Cell cell2 = row.getCell(1);
                String concatenatedValue = "";
                if (cell1 != null) {
                    switch (cell1.getCellType()) {
                        case NUMERIC:
                            concatenatedValue += (int) cell1.getNumericCellValue();
                            break;
                        case STRING:
                            concatenatedValue += cell1.getStringCellValue();
                            break;
                        case FORMULA:
                            switch (formulaEvaluator.evaluate(cell1).getCellType()) {
                                case NUMERIC:
                                    concatenatedValue += (int) formulaEvaluator.evaluate(cell1).getNumberValue();
                                    break;
                                case STRING:
                                    concatenatedValue += formulaEvaluator.evaluate(cell1).getStringValue();
                                    break;
                            }
                            break;
                    }
                }
                if (cell2 != null) {
                    switch (cell2.getCellType()) {
                        case NUMERIC:
                            concatenatedValue += (int) cell2.getNumericCellValue();
                            break;
                        case STRING:
                            concatenatedValue += cell2.getStringCellValue();
                            break;
                        case FORMULA:
                            switch (formulaEvaluator.evaluate(cell2).getCellType()) {
                                case NUMERIC:
                                    concatenatedValue += (int) formulaEvaluator.evaluate(cell2).getNumberValue();
                                    break;
                                case STRING:
                                    concatenatedValue += formulaEvaluator.evaluate(cell2).getStringValue();
                                    break;
                            }
                            break;
                    }
                }
                outputRow.createCell(0).setCellValue(concatenatedValue);
            }

            FileOutputStream outputFile = new FileOutputStream("Lab7-out.xlsx");
            outputWorkbook.write(outputFile);
            outputFile.close();

            file.close();
            inputWorkbook.close();
            outputWorkbook.close();

            System.out.println("Datele au fost scrise cu succes în rezultat.xlsx");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}