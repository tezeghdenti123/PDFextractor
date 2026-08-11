package org.example;


import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import technology.tabula.*;
import technology.tabula.Table;
import technology.tabula.extractors.SpreadsheetExtractionAlgorithm;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class PDFExtractor {
    public static void main(String[] args) {
        String directoryPath = "dec2025"; // 📁 Modifier avec le chemin de ton dossier

        File folder = new File(directoryPath);
        File[] files = folder.listFiles((dir, name) -> name.toLowerCase().endsWith(".pdf"));

        if (files == null || files.length == 0) {
            System.out.println("❌ Aucun fichier PDF trouvé dans le dossier.");
            return;
        }
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Tableau PDF");
        for (File file : files) {
            System.out.println("📂 Traitement du fichier : " + file.getName());
            processPDF(file,workbook,sheet);
        }
        String excelPath = files[0].getParent() + "/" + 1+ ".xlsx";
        FileOutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(excelPath);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }

        System.out.println("✅ Tous les fichiers PDF ont été traités !");
    }

    public static void processPDF(File pdfFile,Workbook workbook,Sheet sheet) {

        try {
            PDDocument document = PDDocument.load(pdfFile);
            ObjectExtractor extractor = new ObjectExtractor(document);
            SpreadsheetExtractionAlgorithm sea = new SpreadsheetExtractionAlgorithm();


            // 🎨 Styles definition (Header + alternating rows)
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setColor(IndexedColors.WHITE.getIndex());
            headerStyle.setFont(headerFont);

            CellStyle whiteRowStyle = workbook.createCellStyle();
            whiteRowStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
            whiteRowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            CellStyle greyRowStyle = workbook.createCellStyle();
            greyRowStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
            greyRowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);



            // Création de l'en-tête
            String[] headers = {"Date", "Date valeur", "Opération", "Débit EUROS", "Crédit EUROS"};
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                cell.setCellStyle(headerStyle);
            }

            List<List<String>> structuredData = new ArrayList<>();
            List<String> currentRow = new ArrayList<>();

            for (int pageNum = 1; pageNum <= document.getNumberOfPages(); pageNum++) {
                Page page = extractor.extract(pageNum);
                List<Table> tables = sea.extract(page);

                for (Table table : tables) {
                    for (List<RectangularTextContainer> row : table.getRows()) {
                        if (row.size() >= 2) { // Vérifier la présence d'une "Date valeur"
                            if (!currentRow.isEmpty() && !row.get(1).getText().isEmpty()) {
                                structuredData.add(new ArrayList<>(currentRow));
                                currentRow.clear();
                            }

                            for (int i = 0; i < row.size(); i++) {
                                if (i < currentRow.size()) {
                                    currentRow.set(i, currentRow.get(i) + " " + row.get(i).getText());
                                } else {
                                    currentRow.add(row.get(i).getText());
                                }
                            }
                        } else {
                            for (int i = 2; i < row.size(); i++) {
                                if (i < currentRow.size()) {
                                    currentRow.set(i, currentRow.get(i) + " " + row.get(i).getText());
                                }
                            }
                        }
                    }
                }
            }

            if (!currentRow.isEmpty()) {
                structuredData.add(currentRow);
            }

            // Écrire dans le fichier Excel avec alternance de couleur et détection d'en-tête PDF
            int rowNum = sheet.getLastRowNum();
            for (List<String> dataRow : structuredData) {
                Row excelRow = sheet.createRow(++rowNum);

                // 🧠 Check if this row looks like a header from the PDF
                boolean isHeaderRow = dataRow.stream().anyMatch(cell ->
                        cell.equalsIgnoreCase("Date") ||
                                cell.equalsIgnoreCase("Date valeur") ||
                                cell.equalsIgnoreCase("Opération") ||
                                cell.toLowerCase().contains("débit") ||
                                cell.toLowerCase().contains("crédit")
                );

                // 🎨 Choose style based on content
                CellStyle rowStyle;
                if (isHeaderRow) {
                    rowStyle = headerStyle; // use blue if PDF row looks like a header
                } else {
                    rowStyle = (rowNum % 2 == 0) ? whiteRowStyle : greyRowStyle;
                }

                // 🧾 Write cells
                for (int i = 0; i < dataRow.size(); i++) {
                    Cell cell = excelRow.createCell(i);
                    cell.setCellValue(dataRow.get(i).trim());
                    cell.setCellStyle(rowStyle);
                }
            }




            document.close();

            System.out.println("✅ Fichier traité avec succès : " );
        } catch (Exception e) {
            System.err.println("❌ Erreur lors du traitement de " + pdfFile.getName() + " : " + e.getMessage());
        }
    }
}
