import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class HtmlToExcel {
    static String singleFilePath = "C:\\Users\\rajub\\AppData\\Roaming\\JetBrains\\IdeaIC2024.2\\scratches\\Test1.html";

    public static void main(String[] args) {
        int totalItemCount = 0;

        // Create a single workbook for merged data
        Workbook mergedWorkbook = new XSSFWorkbook();
        Sheet mergedSheet = mergedWorkbook.createSheet("All Items");

        // Create header row for merged file
        Row mergedHeaderRow = mergedSheet.createRow(0);
        mergedHeaderRow.createCell(0).setCellValue("Item Name");
        mergedHeaderRow.createCell(1).setCellValue("Item Price");
        mergedHeaderRow.createCell(2).setCellValue("Image URL");
        mergedHeaderRow.createCell(3).setCellValue("Item Description");
        mergedHeaderRow.createCell(4).setCellValue("Category");
        mergedHeaderRow.createCell(5).setCellValue("Subcategory");

        int mergedRowIndex = 1;

        Document doc;
        try {
            // Parse the single HTML file
            doc = Jsoup.parse(new File(singleFilePath), "UTF-8");
        } catch (IOException e) {
            System.out.println("Error reading the HTML file: " + e.getMessage());
            return;
        }

        // Select all categories in the single file
        Elements categorySections = doc.select("section.sc-fZEjqe"); // Adjust to match your HTML structure

        for (Element categorySection : categorySections) {
            // Extract category name
            String category = categorySection.select("h4.sc-1hp8d8a-0").text(); // Main category selector

            // Select subcategory sections
            Elements subcategories = categorySection.select("div.sc-izFuNb"); // Subcategory container selector

            // Create individual Excel workbook and sheet for the category
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Items");

            // Create header row
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Item Name");
            headerRow.createCell(1).setCellValue("Item Price");
            headerRow.createCell(2).setCellValue("Image URL");
            headerRow.createCell(3).setCellValue("Item Description");
            headerRow.createCell(4).setCellValue("Category");
            headerRow.createCell(5).setCellValue("Subcategory");

            int rowIndex = 1;
            int categoryItemCount = 0;

            for (Element subCategory : subcategories) {
                String subCategoryName = subCategory.select("p.sc-eEOpXj").text(); // Subcategory name selector

                // Select items under this subcategory
                Elements items = subCategory.select("div.sc-fHeoTs"); // Items selector

                for (Element item : items) {
                    // Extract item details
                    String itemName = item.select("h4.sc-kjKmiN").text(); // Selector for item name
                    String itemPrice = item.select("span.sc-17hyc2s-1").text(); // Selector for item price
                    String itemDescription = item.select("p.sc-kyseJx").text(); // Selector for item description
                    String itemImageUrl = item.select("img.sc-s1isp7-5").attr("src"); // Selector for image URL

                    // Add to individual Excel sheet
                    Row row = sheet.createRow(rowIndex++);
                    row.createCell(0).setCellValue(itemName);
                    row.createCell(1).setCellValue(itemPrice);
                    row.createCell(2).setCellValue(itemImageUrl);
                    row.createCell(3).setCellValue(itemDescription);
                    row.createCell(4).setCellValue(category);
                    row.createCell(5).setCellValue(subCategoryName);

                    // Add the same data to the merged sheet
                    Row mergedRow = mergedSheet.createRow(mergedRowIndex++);
                    mergedRow.createCell(0).setCellValue(itemName);
                    mergedRow.createCell(1).setCellValue(itemPrice);
                    mergedRow.createCell(2).setCellValue(itemImageUrl);
                    mergedRow.createCell(3).setCellValue(itemDescription);
                    mergedRow.createCell(4).setCellValue(category);
                    mergedRow.createCell(5).setCellValue(subCategoryName);

                    categoryItemCount++;
                }
            }

            totalItemCount += categoryItemCount;

            // Write to individual Excel file for the category
            try (FileOutputStream fileOut = new FileOutputStream(category + ".xlsx")) {
                workbook.write(fileOut);
            } catch (IOException e) {
                System.out.println("Error writing the Excel file: " + e.getMessage());
            }

            // Close the individual workbook
            try {
                workbook.close();
            } catch (IOException e) {
                System.out.println("Error closing the workbook: " + e.getMessage());
            }

            System.out.println("Processed " + categoryItemCount + " items across " + subcategories.size() + " subcategories for category: " + category);
        }

        // Write the merged Excel file
        try (FileOutputStream mergedFileOut = new FileOutputStream("Merged_Items.xlsx")) {
            mergedWorkbook.write(mergedFileOut);
        } catch (IOException e) {
            System.out.println("Error writing the merged Excel file: " + e.getMessage());
        }

        // Close the merged workbook
        try {
            mergedWorkbook.close();
        } catch (IOException e) {
            System.out.println("Error closing the merged workbook: " + e.getMessage());
        }

        System.out.println("Total items processed across all categories: " + totalItemCount);
    }
}
