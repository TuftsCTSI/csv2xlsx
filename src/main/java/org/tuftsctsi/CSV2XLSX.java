package org.tuftsctsi;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class CSV2XLSX {

    public static void main(String[] args) {
        try {
            // Create a new workbook
            Workbook workbook = new XSSFWorkbook();

            // Create a sheet in the workbook
            Sheet sheet = workbook.createSheet("Sheet1");

            // Create a row in the sheet
            Row headerRow = sheet.createRow(0);

            // Create cells in the header row
            Cell headerCell1 = headerRow.createCell(0);
            headerCell1.setCellValue("Name");

            Cell headerCell2 = headerRow.createCell(1);
            headerCell2.setCellValue("Age");

            // Create data rows
            Row dataRow1 = sheet.createRow(1);
            dataRow1.createCell(0).setCellValue("John");
            dataRow1.createCell(1).setCellValue(25);

            Row dataRow2 = sheet.createRow(2);
            dataRow2.createCell(0).setCellValue("Jane");
            dataRow2.createCell(1).setCellValue(30);

            // Write the workbook content to a file
            try (FileOutputStream fileOut = new FileOutputStream("workbook.xlsx")) {
                workbook.write(fileOut);
                System.out.println("Excel file created successfully!");
            }

            // Close the workbook
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
