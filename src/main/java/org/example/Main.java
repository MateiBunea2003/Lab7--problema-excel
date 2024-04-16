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
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

    public static void main(String[] args) {

        try {
            FileInputStream fileInput = new FileInputStream(new File("Lab7.xlsx"));
            XSSFWorkbook workbookInput = new XSSFWorkbook(fileInput);
            XSSFSheet sheetInput = workbookInput.getSheetAt(0);

            // Crearea listei pentru elementele de adăugat în a treia coloană
            List<String> elements = new ArrayList<>();
            elements.add("1A");
            elements.add("2B");
            elements.add("3C");
            elements.add("6");

            FileInputStream fileOutput = new FileInputStream(new File("Lab7-out.xlsx"));
            XSSFWorkbook workbookOutput = new XSSFWorkbook(fileOutput);
            XSSFSheet sheetOutput = workbookOutput.getSheetAt(0);

            Iterator<Row> rowIterator = sheetInput.iterator();
            int rowIndexOutput = 0;

            while (rowIterator.hasNext()) {
                Row rowInput = rowIterator.next();
                Row rowOutput = sheetOutput.createRow(rowIndexOutput++);

                Iterator<Cell> cellIterator = rowInput.cellIterator();
                int cellIndex = 0;

                while (cellIterator.hasNext()) {
                    Cell cellInput = cellIterator.next();
                    Cell cellOutput = rowOutput.createCell(cellIndex++);

                    switch (cellInput.getCellType()) {
                        case NUMERIC:
                            cellOutput.setCellValue(cellInput.getNumericCellValue());
                            break;
                        case STRING:
                            cellOutput.setCellValue(cellInput.getStringCellValue());
                            break;
                        case FORMULA:
                            cellOutput.setCellValue(cellInput.getCellFormula());
                            break;
                        default:
                            break;
                    }
                }

                // Adăugarea elementelor în a treia coloană
                Cell cellOutput = rowOutput.createCell(cellIndex);
                if (!elements.isEmpty()) {
                    String element = elements.remove(0);
                    cellOutput.setCellValue(element);
                }
            }

            fileInput.close();
            fileOutput.close();

            FileOutputStream outputStream = new FileOutputStream("Lab7-out.xlsx");
            workbookOutput.write(outputStream);
            workbookOutput.close();
            outputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
