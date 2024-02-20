package org.example;
import org.apache.poi.ss.usermodel.*;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class SeleniumTest {
    public static void main(String[] args) throws InterruptedException, IOException {
        FileInputStream fileInputStream = new FileInputStream("C://Users//nishat//Desktop//login.xlsx");
        Workbook workbook = WorkbookFactory.create(fileInputStream);
        Sheet sheet = workbook.getSheet("credentials");
        Sheet urlSheet = workbook.getSheet("Second_website");
        List<String> driveLinks = new ArrayList<>();
        int urlRowCount = urlSheet.getLastRowNum();
        int urlColCount = urlSheet.getRow(0).getLastCellNum();

        for (int rowIndex = 0; rowIndex <= urlRowCount; rowIndex++) {
            Row row = urlSheet.getRow(rowIndex);

            // Get the value from column 1
            String drivelink = row.getCell(1).getStringCellValue();
            System.out.println("Drive Link: " + drivelink);
            driveLinks.add(drivelink);
            // Use the value as a variable in your main loop
        }
        System.out.println(driveLinks.get(0));


        System.setProperty("webdriver.chrome.driver",driveLinks.get(0));
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();

        WebDriverWait wait = new WebDriverWait(driver, 100);



        // Iterate through rows (including the header row)
        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);

            // Get username and password from Excel
            String username = row.getCell(0).getStringCellValue();
            String password = row.getCell(1).getStringCellValue();

            // Navigate to the login page
            driver.get(driveLinks.get(1));
            Thread.sleep(500);
            driver.findElement(By.xpath(driveLinks.get(2))).click();


            // Find username and password fields and enter values
            driver.findElement(By.xpath(driveLinks.get(3))).sendKeys(username);
            driver.findElement(By.xpath(driveLinks.get(4))).sendKeys(password);
            Thread.sleep(500);
            // Click the login button
            driver.findElement(By.xpath(driveLinks.get(5))).click();
            Thread.sleep(1500);

            // Check the current URL
            String cururl = driveLinks.get(6);
            String currentUrl = driver.getCurrentUrl();
            System.out.println(currentUrl);

            // Update Excel based on the test result
            Cell resultCell = row.createCell(2); // Assuming the result column is the third column
            if (currentUrl.equals(cururl)) {
                System.out.println("Test case passed");
                resultCell.setCellValue("Pass");
                driver.findElement(By.xpath(driveLinks.get(7))).click(); Thread.sleep(300);
                driver.findElement(By.xpath(driveLinks.get(8))).click();
                //driver.get(driveLinks.get(1));
                Thread.sleep(500);

                driver.navigate().refresh();
            } else {
                System.out.println("Test case failed");
                resultCell.setCellValue("Fail");
                Thread.sleep(400);
                driver.navigate().refresh();
            }
        }

        // Save the results to the output file after the loop
        FileOutputStream fileOutputStream = new FileOutputStream("C://Users//nishat//Desktop//login.xlsx");
        workbook.write(fileOutputStream);

        // Close resources
        fileInputStream.close();
        fileOutputStream.close();
        workbook.close();

        // Quit the WebDriver
        driver.quit();
    }
}