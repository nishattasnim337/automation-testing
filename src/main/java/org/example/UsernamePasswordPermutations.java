package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import static org.example.HRM_Login.generateLimitedPermutations;

public class UsernamePasswordPermutations {
    private static final int MAX_PERMUTATIONS = 15;
    public static void main(String[] args) throws IOException {
        // Specify the path to the Excel file
        String excelFilePath = "C://Users//nishat//Desktop//login.xlsx";

        // Read usernames and passwords from Excel file
        List<String> usernames = readColumnFromExcel(excelFilePath, "login", 0);
        List<String> passwords = readColumnFromExcel(excelFilePath, "login", 1);

        // Log the input data
        System.out.println("Usernames: " + usernames);
        System.out.println("Passwords: " + passwords);

        // Open the existing workbook
        try (FileInputStream file = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(file)) {

            // Create a new sheet named "credentials" or get the existing one
            Sheet credentialsSheet = workbook.getSheet("credentials");
            if (credentialsSheet == null) {
                credentialsSheet = workbook.createSheet("credentials");
                // Write headers if the sheet is newly created
                Row headerRow = credentialsSheet.createRow(0);
                headerRow.createCell(0).setCellValue("Username");
                headerRow.createCell(1).setCellValue("Password");
            } else {
                // Clear existing data from the sheet
                clearSheet(credentialsSheet);
            }

            // Iterate over each username and password
            int maxSize = Math.min(usernames.size(), passwords.size());
            for (int i = 0; i < maxSize; i++) {
                String username = usernames.get(i);
                String password = passwords.get(i);

                // Log the current username and password
                System.out.println("Processing: Username=" + username + ", Password=" + password);

                // Check if the username contains '@' sign
                int atIndex = username.indexOf('@');
                if (atIndex != -1) {
                    // If '@' sign is present, permute only the part before '@'
                    String usernameBeforeAt = username.substring(0, atIndex);
                    List<String> usernamePermutations = generateLimitedPermutations(usernameBeforeAt, 6);
                    // Append the part after '@' unchanged
                    usernamePermutations.replaceAll(permutation -> permutation + username.substring(atIndex));
                    // Password permutations remain unchanged
                    List<String> passwordPermutations = generateLimitedPermutations(password, 6);

                    // Write permutated data to the sheet
                    for (int j = 0; j < usernamePermutations.size(); j++) {
                        Row newRow = credentialsSheet.createRow(credentialsSheet.getLastRowNum() + 1);
                        newRow.createCell(0).setCellValue(usernamePermutations.get(j));
                        newRow.createCell(1).setCellValue(passwordPermutations.get(j));
                    }
                } else {
                    // If no '@' sign is present, permute the entire username
                    List<String> usernamePermutations = generateLimitedPermutations(username, 6);
                    List<String> passwordPermutations = generateLimitedPermutations(password, 6);

                    // Write permutated data to the sheet
                    for (int j = 0; j < usernamePermutations.size(); j++) {
                        Row newRow = credentialsSheet.createRow(credentialsSheet.getLastRowNum() + 1);
                        newRow.createCell(0).setCellValue(usernamePermutations.get(j));
                        newRow.createCell(1).setCellValue(passwordPermutations.get(j));
                    }
                }
            }

            // Write the changes back to the existing Excel file
            try (FileOutputStream outFile = new FileOutputStream(excelFilePath)) {
                workbook.write(outFile);
                System.out.println("Permutated data written to the 'credentials' sheet successfully.");
            }
        }
    }

    private static void clearSheet(Sheet sheet) {
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i <= lastRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                sheet.removeRow(row);
            }
        }
    }

    private static List<String> readColumnFromExcel(String filePath, String sheetName, int columnIndex) throws IOException {
        List<String> data = new ArrayList<>();

        try (FileInputStream file = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(file)) {
            Sheet sheet = workbook.getSheet(sheetName);

            Iterator<Row> iterator = sheet.iterator();

            while (iterator.hasNext()) {
                Row currentRow = iterator.next();
                Cell cell = currentRow.getCell(columnIndex);

                if (cell != null) {
                    if (cell.getCellType() == CellType.NUMERIC) {
                        // Handle numeric values as integers
                        data.add(String.valueOf((int) cell.getNumericCellValue()));
                    } else {
                        data.add(cell.toString());
                    }
                }
            }
        }

        return data;
    }
}