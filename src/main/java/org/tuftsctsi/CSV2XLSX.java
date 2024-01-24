package org.tuftsctsi;

import java.io.*;
import java.security.GeneralSecurityException;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.poifs.crypt.*;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CSV2XLSX {

    public static void main(String[] args) {
        try {

            // Get password as first argument
            if (args.length < 1) {
                System.out.println("Please provide a password as a command-line argument.");
                return;
            }

            String password = args[0];

            // Read CSV data from stdin
            List<List<String>> csvData = readCSV(System.in);

            // Create a new XLSX workbook
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Sheet1");

            // Write CSV data to XLSX
            writeDataToSheet(csvData, sheet);


            // Save the workbook to an encrypted file
            encryptAndSaveWorkbook(workbook, password, "output.xlsx");

            System.out.println("Conversion successful!");

        } catch (Exception e) {
            System.out.println("Exception while writting protected xlsx file");
            e.printStackTrace();
        }
    }

    private static void encryptAndSaveWorkbook(Workbook workbook, String password, String filePath)
            throws InvalidFormatException, IOException, GeneralSecurityException {

        try (POIFSFileSystem fs = new POIFSFileSystem()) {

            File file = new File(filePath);
            FileOutputStream fos = new FileOutputStream(file);
            workbook.write(fos);
            fos.close();

            EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
            Encryptor encryptor = info.getEncryptor();
            encryptor.confirmPassword(password);
            
            try (OPCPackage opc = OPCPackage.open(file, PackageAccess.READ_WRITE);
                 OutputStream os = encryptor.getDataStream(fs)) {
                 opc.save(os);
            }

            try (FileOutputStream efos = new FileOutputStream(filePath)) {
                fs.writeFilesystem(efos);
            }

        } catch (InvalidFormatException | IOException | GeneralSecurityException e) {
            System.out.println("Exception while writting protected xlsx file");
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
