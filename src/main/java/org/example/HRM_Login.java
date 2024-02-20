package org.example;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;


public class HRM_Login {

    public static void main(String[] args) throws IOException, InterruptedException {

        FileInputStream file = new FileInputStream("C://Users//nishat//Desktop//login.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet loginSheet = workbook.getSheet("login");

        int loginRowCount = loginSheet.getLastRowNum();
        int loginColCount = loginSheet.getRow(0).getLastCellNum();

        XSSFSheet credentialSheet = workbook.getSheet("credentials");

        // If 'credentials' sheet doesn't exist, create it
        if (credentialSheet == null) {
            credentialSheet = workbook.createSheet("credentials");
        }

        for (int i = 1; i <= loginRowCount; i++) {
            XSSFRow currentRow = loginSheet.getRow(i);
            String username = currentRow.getCell(0).toString();
            XSSFCell passCell = currentRow.getCell(1);
            int password = (int) passCell.getNumericCellValue();

            // Generate 5 new username and password combinations for each original username and password
            List<String> usernamePermutations = generateLimitedPermutations(username, 6);
            List<String> passwordPermutations = generateLimitedPermutations(String.valueOf(password), 6);

            for (int j = 0; j < 5; j++) {
                String newUsername = usernamePermutations.get(j);
                int newPassword = Integer.parseInt(passwordPermutations.get(j));

                XSSFRow newRow = credentialSheet.createRow((i - 1) * 5 + j);
                newRow.createCell(0).setCellValue(newUsername);
                newRow.createCell(1).setCellValue(newPassword);
            }
        }

        // Write the changes back to the workbook
        try (FileOutputStream outFile = new FileOutputStream("C://Users//nishat//Desktop//login.xlsx")) {
            workbook.write(outFile);
        }

        System.out.println("Permutations generated and written to credentials sheet successfully.");
    }

    public static List<String> generateLimitedPermutations(String input, int limit) {
        List<String> result = new ArrayList<>();
        char[] characters = input.toCharArray();
        boolean[] used = new boolean[characters.length];
        StringBuilder currentPermutation = new StringBuilder();

        generateLimitedPermutationsHelper(characters, used, currentPermutation, result, limit);

        return result;
    }

    private static void generateLimitedPermutationsHelper(char[] characters, boolean[] used, StringBuilder currentPermutation, List<String> result, int limit) {
        if (result.size() == limit) {
            return;
        }

        if (currentPermutation.length() == characters.length) {
            result.add(currentPermutation.toString());
            return;
        }

        for (int i = 0; i < characters.length; i++) {
            if (!used[i]) {
                used[i] = true;
                currentPermutation.append(characters[i]);

                generateLimitedPermutationsHelper(characters, used, currentPermutation, result, limit);

                // Backtrack
                used[i] = false;
                currentPermutation.deleteCharAt(currentPermutation.length() - 1);
            }
        }


//        FileInputStream file = new FileInputStream("C://Users//nishat//Desktop//login.xlsx");
//        XSSFWorkbook workbook = new XSSFWorkbook(file);
//        XSSFSheet loginSheet = workbook.getSheet("login");
//
//        int loginRowCount = loginSheet.getLastRowNum();
//        int loginColCount = loginSheet.getRow(0).getLastCellNum();
//
//        XSSFSheet credentialSheet = workbook.getSheet("credentials");
//
//        // If 'credentials' sheet doesn't exist, create it
//        if (credentialSheet == null) {
//            credentialSheet = workbook.createSheet("credentials");
//        }
//
//        for (int i = 1; i <= loginRowCount; i++) {
//            XSSFRow currentRow = loginSheet.getRow(i);
//            String username = currentRow.getCell(0).toString();
//            XSSFCell passCell = currentRow.getCell(1);
//            int password = (int) passCell.getNumericCellValue();
//
//            // Create 10 new username and password combinations for each original username and password
//            for (int j = 0; j < 10; j++) {
//
//                // add permutation and combination
//
//                String newUsername = username + "_perm" + j;
//                int newPassword = password + j;
//
//                XSSFRow newRow = credentialSheet.createRow((i - 1) * 10 + j);
//                newRow.createCell(0).setCellValue(newUsername);
//                newRow.createCell(1).setCellValue(newPassword);
//            }
//        }
//
//        // Write the changes back to the workbook
//        try (FileOutputStream outFile = new FileOutputStream("C://Users//nishat//Desktop//login.xlsx")) {
//            workbook.write(outFile);
//        }
//
//        System.out.println("Permutations generated and written to credentials sheet successfully.");
    }}
        // Close the workbo

//        System.out.println("Hello");

//        FileInputStream file = new FileInputStream("C://Users//nishat//Desktop//login.xlsx");
//        XSSFWorkbook workbook = new XSSFWorkbook(file);
//        XSSFSheet sheet = workbook.getSheet("login");
//        //XSSFSheet sheet2 = workbook.getSheet("credential");
//        System.out.println("Hello");
//        int rowcount = sheet.getLastRowNum();
//        int colcount = sheet.getRow(0).getLastCellNum();
//        System.out.println(rowcount);
//        System.out.println(colcount);
//        for (int i = 1; i <= rowcount; i++) {
//           // System.out.println("hello");
//            //XSSFRow outputrow=outputsheet.createRow(i);
//            XSSFRow currentrow = sheet.getRow(i);
//            String name = currentrow.getCell(0).toString();
//            XSSFCell pass = currentrow.getCell(1);
//            int password = (int) pass.getNumericCellValue();
//            System.out.println(name);
//            System.out.println(password);







