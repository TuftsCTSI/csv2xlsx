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

            // Parse command line arguments
            String filePath = null;
            String password = null;

            for (int i = 0; i < args.length; i += 2) {
                switch (args[i]) {
                    case "--file":
                        filePath = args[i + 1];
                        break;
                    case "--password":
                        password = args[i + 1];
                        break;
                    default:
                        System.out.println("Invalid argument: " + args[i]);
                        System.exit(1);
                }
            }

            // Check if both file name and password are provided
            if (filePath == null || password == null) {
                System.out.println("Both file name and password are required.");
                System.exit(1);
            }

            processFile(filePath, password);

        } catch (Exception e) {
            System.out.println("Exception while writting protected xlsx file");
            e.printStackTrace();
        }
    }

    private static void processFile(String filePath, String password) throws Exception {
            List<List<String>> csvData = readCSV(System.in);
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Sheet1");
            writeDataToSheet(csvData, sheet);
            encryptAndSaveWorkbook(workbook, filePath, password);
    }

    private static void encryptAndSaveWorkbook(Workbook workbook, String filePath, String password)
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
