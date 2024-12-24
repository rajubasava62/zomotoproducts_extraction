import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class Dmart {
    public static void main(String[] args) {
        try {
            // Path to your HTML file in the IntelliJ Scratch folder
            String filePath = "C:\\Users\\rajub\\AppData\\Roaming\\JetBrains\\IdeaIC2024.2\\scratches\\dmart.html";

            // Read the HTML file content
            String html = new String(Files.readAllBytes(Paths.get(filePath)));

            // Parse the HTML
            Document document = Jsoup.parse(html);

            // Find all product cards
            Elements productCards = document.select(".horizontal-card_card-horizontal__nJVj3");

            // Create an Excel workbook and sheet
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Products");

            // Add column headers
            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Product Name");
            headerRow.createCell(1).setCellValue("MRP");
            headerRow.createCell(2).setCellValue("DMart Price");
            headerRow.createCell(3).setCellValue("Quantity");
            headerRow.createCell(4).setCellValue("Image URL");

            // Process each product card
            int rowIndex = 1;
            for (Element card : productCards) {
                Row row = sheet.createRow(rowIndex++);

                // Extract data
                String productName = card.select(".horizontal-card_title__vAl_A").text();
                String mrp = card.select(".horizontal-card_section-left__BJFkt .horizontal-card_strike-through__YIsFD .horizontal-card_amount__4vA7J").text();
                String dmartPrice = card.select(".horizontal-card_section-left__BJFkt .horizontal-card_price-container__yPWSi:last-child .horizontal-card_amount__4vA7J").text();
                String quantity = card.select(".bootstrap-select_option-mobile__qXEdO span:first-child").text();

                String imageUrl = card.select(".horizontal-card_image__LJJTX").attr("style");
                String finalImageUrl = "Image URL not found"; // Default fallback

                Pattern pattern = Pattern.compile("background-image:\\s*url\\(\"(.*?)\"\\)");


              Matcher matcher = pattern.matcher(imageUrl);

                if (matcher.find()) {
                    // Extract the first valid image URL
                    finalImageUrl = matcher.group(1);
                }
// Assign the extracted URL to finalImageUrl
                imageUrl = finalImageUrl;



                // Write data to Excel row
                row.createCell(0).setCellValue(productName);
                row.createCell(1).setCellValue(mrp.replace("₹", "").trim());
                row.createCell(2).setCellValue(dmartPrice.replace("₹", "").trim());

                row.createCell(3).setCellValue(quantity);
                row.createCell(4).setCellValue(imageUrl);
            }

            // Write the Excel file
            try (FileOutputStream fileOut = new FileOutputStream("D-Products.xlsx")) {
                workbook.write(fileOut);
            }

            System.out.println("Data has been extracted to Products.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
