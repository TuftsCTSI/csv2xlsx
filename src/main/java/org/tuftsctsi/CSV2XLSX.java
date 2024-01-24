package org.tuftsctsi;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CSV2XLSX {

    public static void main(String[] args) {
        try {
            // Read CSV data from stdin
            List<List<String>> csvData = readCSV(System.in);

            // Create a new XLSX workbook
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Sheet1");

            // Write CSV data to XLSX
            writeDataToSheet(csvData, sheet);

            // Save the workbook to a file or do other operations as needed
            // For example, you can save it to a file using:
            FileOutputStream outputStream = new FileOutputStream("output.xlsx");
            workbook.write(outputStream);
            outputStream.close();

            System.out.println("Conversion successful!");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static List<List<String>> readCSV(InputStream inputStream) throws IOException {
        BufferedReader reader = new BufferedReader(new InputStreamReader(inputStream));
        String line;
        List<List<String>> data = new java.util.ArrayList<>();

        while ((line = reader.readLine()) != null) {
            List<String> values = Arrays.asList(line.split(","));
            data.add(values);
        }

        return data;
    }

    private static void writeDataToSheet(List<List<String>> data, Sheet sheet) {
        int rowIndex = 0;
        for (List<String> rowData : data) {
            Row row = sheet.createRow(rowIndex++);
            int cellIndex = 0;
            for (String cellData : rowData) {
                Cell cell = row.createCell(cellIndex++);
                cell.setCellValue(cellData);
            }
        }
    }
}
