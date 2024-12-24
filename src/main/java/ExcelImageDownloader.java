import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.net.URL;

public class ExcelImageDownloader {
    public static void main(String[] args) {
        // Specify the path to your Excel file
        String excelFilePath;
        if (args.length < 1) {
            excelFilePath = "C:\\Users\\rajub\\IdeaProjects\\D-mart\\image.xlsx"; // Update this path
        } else {
            excelFilePath = args[0];
        }

        // Read the Excel file
        try (FileInputStream fileInputStream = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fileInputStream)) {

            Sheet sheet = workbook.getSheetAt(0); // Read the first sheet
            Row headerRow = sheet.getRow(0);

            if (headerRow == null) {
                System.out.println("The Excel file is empty or missing headers.");
                return;
            }

            // Identify the column indexes for the required headers
            int itemImageIndex = -1;
            int itemNameIndex = -1;

            for (Cell cell : headerRow) {
                String headerValue = cell.getStringCellValue().trim().toLowerCase();
                if (headerValue.equals("image_url")) {
                    itemImageIndex = cell.getColumnIndex();
                } else if (headerValue.equals("image_name")) {
                    itemNameIndex = cell.getColumnIndex();
                }
            }

            if (itemImageIndex == -1 || itemNameIndex == -1) {
                System.out.println("The required columns ('itemimage' or 'itemname') are missing in the Excel file.");
                return;
            }

            // Iterate through rows and process images
            for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) continue;

                Cell imageCell = row.getCell(itemImageIndex);
                Cell nameCell = row.getCell(itemNameIndex);

                if (imageCell != null && imageCell.getCellType() == CellType.STRING
                        && nameCell != null && nameCell.getCellType() == CellType.STRING) {
                    String imageUrl = imageCell.getStringCellValue();
                    String itemName = nameCell.getStringCellValue().replaceAll("[\\/:*?\"<>|]", "_"); // Sanitize file name

                    if (!imageUrl.isEmpty()) {
                        downloadImage(imageUrl, itemName + ".jpg");
                    }
                }
            }

        } catch (IOException e) {
            System.out.println("Error reading the Excel file: " + e.getMessage());
        }
    }

    private static void downloadImage(String imageUrl, String outputFileName) {
        String saveDirectory = "C:\\Users\\rajub\\OneDrive\\Desktop\\New folder\\"; // Set your desired directory
        new File(saveDirectory).mkdirs(); // Create the directory if it doesn't exist

        try (InputStream in = new URL(imageUrl).openStream();
             FileOutputStream out = new FileOutputStream(saveDirectory + outputFileName)) {

            byte[] buffer = new byte[1024];
            int bytesRead;
            while ((bytesRead = in.read(buffer)) != -1) {
                out.write(buffer, 0, bytesRead);
            }

            System.out.println("Downloaded image: " + saveDirectory + outputFileName);

        } catch (IOException e) {
            System.out.println("Error downloading image from URL " + imageUrl + ": " + e.getMessage());
        }
    }
}
