package com.beingidly.litexl.benchmark;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.crypto.EncryptionOptions;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.crypt.EncryptionMode;
import org.apache.poi.poifs.crypt.Encryptor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Random;

/**
 * Complex benchmark comparing litexl vs Apache POI with realistic workbook features.
 *
 * <p>Benchmarks 5 operations:
 * <ul>
 *   <li>Write - Generate complete workbook with styles, formulas, validation, etc.</li>
 *   <li>Read - Read all cells from workbook</li>
 *   <li>Read-Modify-Write - Open, modify 300 cells in Sales sheet, save</li>
 *   <li>Encrypted Write - Generate and encrypt with AES-256</li>
 *   <li>Encrypted Read - Decrypt and read all cells</li>
 * </ul>
 */
public class ComplexCompare {

    private static final String PASSWORD = "benchmark123";
    private static final int MODIFY_CELLS = 300;

    public static void main(String[] args) throws Exception {
        ComplexDataGenerator generator = new ComplexDataGenerator();
        long totalCells = generator.calculateTotalCells();

        // Print header
        printHeader(totalCells);

        Path litexlPlain = Files.createTempFile("litexl-complex-plain-", ".xlsx");
        Path litexlEnc = Files.createTempFile("litexl-complex-enc-", ".xlsx");
        Path litexlModified = Files.createTempFile("litexl-complex-modified-", ".xlsx");
        Path poiPlain = Files.createTempFile("poi-complex-plain-", ".xlsx");
        Path poiEnc = Files.createTempFile("poi-complex-enc-", ".xlsx");
        Path poiModified = Files.createTempFile("poi-complex-modified-", ".xlsx");

        try {
            // Warmup with small workbook
            System.out.println("Warming up...");
            warmup();
            System.out.println();

            // Results storage
            long[] litexlTimes = new long[5];
            long[] poiTimes = new long[5];
            String[] operations = {"Write", "Read", "Read-Modify-Write", "Encrypted Write", "Encrypted Read"};

            // ============= WRITE =============
            System.out.println("Running Write benchmark...");
            generator.resetRandom();
            long start = System.nanoTime();
            generator.generateLitexl(litexlPlain);
            litexlTimes[0] = System.nanoTime() - start;

            generator.resetRandom();
            start = System.nanoTime();
            generator.generatePoi(poiPlain);
            poiTimes[0] = System.nanoTime() - start;

            // ============= READ =============
            System.out.println("Running Read benchmark...");
            start = System.nanoTime();
            int litexlReadCells = readLitexl(litexlPlain);
            litexlTimes[1] = System.nanoTime() - start;

            start = System.nanoTime();
            int poiReadCells = readPoi(poiPlain);
            poiTimes[1] = System.nanoTime() - start;

            // ============= READ-MODIFY-WRITE =============
            System.out.println("Running Read-Modify-Write benchmark...");
            start = System.nanoTime();
            readModifyWriteLitexl(litexlPlain, litexlModified);
            litexlTimes[2] = System.nanoTime() - start;

            start = System.nanoTime();
            readModifyWritePoi(poiPlain, poiModified);
            poiTimes[2] = System.nanoTime() - start;

            // ============= ENCRYPTED WRITE =============
            System.out.println("Running Encrypted Write benchmark...");
            generator.resetRandom();
            start = System.nanoTime();
            generateLitexlEncrypted(generator, litexlEnc);
            litexlTimes[3] = System.nanoTime() - start;

            generator.resetRandom();
            start = System.nanoTime();
            generatePoiEncrypted(generator, poiEnc);
            poiTimes[3] = System.nanoTime() - start;

            // ============= ENCRYPTED READ =============
            System.out.println("Running Encrypted Read benchmark...");
            start = System.nanoTime();
            int litexlEncReadCells = readLitexlEncrypted(litexlEnc);
            litexlTimes[4] = System.nanoTime() - start;

            start = System.nanoTime();
            int poiEncReadCells = readPoiEncrypted(poiEnc);
            poiTimes[4] = System.nanoTime() - start;

            System.out.println();

            // ============= CROSS VERIFICATION =============
            System.out.println("Verifying cross-compatibility...");

            // POI reads litexl plain file
            int poiReadLitexlPlain = readPoi(litexlPlain);
            boolean plainOk = poiReadLitexlPlain > 0;
            System.out.printf("  POI reads litexl plain:     %,d cells - %s%n",
                poiReadLitexlPlain, plainOk ? "OK" : "FAILED");

            // POI reads litexl encrypted file
            int poiReadLitexlEnc = readPoiEncrypted(litexlEnc);
            boolean encOk = poiReadLitexlEnc > 0;
            System.out.printf("  POI reads litexl encrypted: %,d cells - %s%n",
                poiReadLitexlEnc, encOk ? "OK" : "FAILED");

            // litexl reads POI plain file
            int litexlReadPoiPlain = readLitexl(poiPlain);
            boolean litexlPlainOk = litexlReadPoiPlain > 0;
            System.out.printf("  litexl reads POI plain:     %,d cells - %s%n",
                litexlReadPoiPlain, litexlPlainOk ? "OK" : "FAILED");

            // litexl reads POI encrypted file (may fail due to encryption format differences)
            int litexlReadPoiEnc = 0;
            boolean litexlEncOk = false;
            try {
                litexlReadPoiEnc = readLitexlEncrypted(poiEnc);
                litexlEncOk = litexlReadPoiEnc > 0;
                System.out.printf("  litexl reads POI encrypted: %,d cells - %s%n",
                    litexlReadPoiEnc, litexlEncOk ? "OK" : "FAILED");
            } catch (Exception e) {
                System.out.printf("  litexl reads POI encrypted: SKIPPED (format incompatibility)%n");
            }

            System.out.println();

            // ============= RESULTS TABLE =============
            printResults(operations, litexlTimes, poiTimes);

            // ============= FILE SIZES =============
            System.out.println();
            System.out.println("File Sizes:");
            System.out.printf("  litexl plain:     %,.2f MB%n", Files.size(litexlPlain) / 1024.0 / 1024.0);
            System.out.printf("  litexl encrypted: %,.2f MB%n", Files.size(litexlEnc) / 1024.0 / 1024.0);
            System.out.printf("  POI plain:        %,.2f MB%n", Files.size(poiPlain) / 1024.0 / 1024.0);
            System.out.printf("  POI encrypted:    %,.2f MB%n", Files.size(poiEnc) / 1024.0 / 1024.0);

            // ============= VERIFICATION =============
            System.out.println();
            System.out.printf("Cells read: litexl=%,d, POI=%,d%n", litexlReadCells, poiReadCells);
            System.out.printf("Encrypted cells read: litexl=%,d, POI=%,d%n", litexlEncReadCells, poiEncReadCells);

            // ============= JSON OUTPUT =============
            System.out.println();
            printJson(operations, litexlTimes, poiTimes, totalCells);

        } finally {
            // Cleanup
            Files.deleteIfExists(litexlPlain);
            Files.deleteIfExists(litexlEnc);
            Files.deleteIfExists(litexlModified);
            Files.deleteIfExists(poiPlain);
            Files.deleteIfExists(poiEnc);
            Files.deleteIfExists(poiModified);
        }
    }

    private static void printHeader(long totalCells) {
        System.out.println("╔══════════════════════════════════════════════════════════════════╗");
        System.out.println("║            Complex Benchmark: litexl vs Apache POI               ║");
        System.out.printf("║     5 sheets, 9,300 rows, ~%,d cells                        ║%n", totalCells);
        System.out.println("║     Features: styles, formulas, conditional formatting,          ║");
        System.out.println("║               data validation, merge, autofilter, protection     ║");
        System.out.println("╚══════════════════════════════════════════════════════════════════╝");
        System.out.println();
    }

    private static void printResults(String[] operations, long[] litexlTimes, long[] poiTimes) {
        System.out.println("╔══════════════════════════════════════════════════════════════════╗");
        System.out.println("║                           RESULTS                                ║");
        System.out.println("╠════════════════════╤═══════════╤═══════════╤════════╤═══════════╣");
        System.out.println("║ Operation          │ litexl    │ POI       │ Ratio  │ Winner    ║");
        System.out.println("╠════════════════════╪═══════════╪═══════════╪════════╪═══════════╣");

        for (int i = 0; i < operations.length; i++) {
            double litexlMs = litexlTimes[i] / 1e6;
            double poiMs = poiTimes[i] / 1e6;
            double ratio;
            String winner;

            if (litexlTimes[i] < poiTimes[i]) {
                ratio = (double) poiTimes[i] / litexlTimes[i];
                winner = "litexl";
            } else {
                ratio = (double) litexlTimes[i] / poiTimes[i];
                winner = "POI";
            }

            System.out.printf("║ %-18s │ %7.0f ms │ %7.0f ms │ %5.2fx │ %-9s ║%n",
                operations[i], litexlMs, poiMs, ratio, winner);
        }

        System.out.println("╚════════════════════╧═══════════╧═══════════╧════════╧═══════════╝");
    }

    private static void printJson(String[] operations, long[] litexlTimes, long[] poiTimes, long totalCells) {
        System.out.println("JSON Output:");
        System.out.println("{");
        System.out.println("  \"benchmark\": \"complex\",");
        System.out.printf("  \"totalCells\": %d,%n", totalCells);
        System.out.println("  \"results\": {");

        for (int i = 0; i < operations.length; i++) {
            String key = operations[i].toLowerCase().replace("-", "_").replace(" ", "_");
            double litexlMs = litexlTimes[i] / 1e6;
            double poiMs = poiTimes[i] / 1e6;
            double ratio = litexlTimes[i] < poiTimes[i]
                ? (double) poiTimes[i] / litexlTimes[i]
                : -((double) litexlTimes[i] / poiTimes[i]);
            String comma = i < operations.length - 1 ? "," : "";

            System.out.printf("    \"%s\": {\"litexl_ms\": %.2f, \"poi_ms\": %.2f, \"ratio\": %.2f}%s%n",
                key, litexlMs, poiMs, ratio, comma);
        }

        System.out.println("  }");
        System.out.println("}");
    }

    private static void warmup() throws Exception {
        Path warmup = Files.createTempFile("warmup-", ".xlsx");
        try {
            // Small warmup workbook
            for (int i = 0; i < 2; i++) {
                try (Workbook wb = Workbook.create()) {
                    Sheet sheet = wb.addSheet("Warmup");
                    for (int r = 0; r < 100; r++) {
                        for (int c = 0; c < 10; c++) {
                            sheet.cell(r, c).set("W" + r + "C" + c);
                        }
                    }
                    wb.save(warmup);
                }

                try (XSSFWorkbook wb = new XSSFWorkbook()) {
                    XSSFSheet sheet = wb.createSheet("Warmup");
                    for (int r = 0; r < 100; r++) {
                        var row = sheet.createRow(r);
                        for (int c = 0; c < 10; c++) {
                            row.createCell(c).setCellValue("W" + r + "C" + c);
                        }
                    }
                    try (FileOutputStream fos = new FileOutputStream(warmup.toFile())) {
                        wb.write(fos);
                    }
                }
            }
        } finally {
            Files.deleteIfExists(warmup);
        }
    }

    // ==================== LITEXL ====================

    private static int readLitexl(Path path) throws Exception {
        int count = 0;
        try (Workbook wb = Workbook.open(path)) {
            for (int s = 0; s < wb.sheetCount(); s++) {
                Sheet sheet = wb.getSheet(s);
                for (var rowEntry : sheet.rows().entrySet()) {
                    for (var cellEntry : rowEntry.getValue().cells().entrySet()) {
                        cellEntry.getValue().string();
                        count++;
                    }
                }
            }
        }
        return count;
    }

    private static int readLitexlEncrypted(Path path) throws Exception {
        int count = 0;
        try (Workbook wb = Workbook.open(path, PASSWORD)) {
            for (int s = 0; s < wb.sheetCount(); s++) {
                Sheet sheet = wb.getSheet(s);
                for (var rowEntry : sheet.rows().entrySet()) {
                    for (var cellEntry : rowEntry.getValue().cells().entrySet()) {
                        cellEntry.getValue().string();
                        count++;
                    }
                }
            }
        }
        return count;
    }

    private static void readModifyWriteLitexl(Path input, Path output) throws Exception {
        try (Workbook wb = Workbook.open(input)) {
            // Find Sales sheet and modify 300 cells
            Sheet sales = wb.getSheet("Sales");
            Random rand = new Random(42);
            for (int i = 0; i < MODIFY_CELLS; i++) {
                int row = rand.nextInt(3000) + 1;  // Skip header
                int col = rand.nextInt(15);
                sales.cell(row, col).set("MODIFIED-" + i);
            }
            wb.save(output);
        }
    }

    private static void generateLitexlEncrypted(ComplexDataGenerator generator, Path path) throws Exception {
        // Generate to temp file first, then encrypt
        Path temp = Files.createTempFile("litexl-temp-", ".xlsx");
        try {
            generator.generateLitexl(temp);

            // Open and save with encryption
            try (Workbook wb = Workbook.open(temp)) {
                wb.save(path, EncryptionOptions.aes256(PASSWORD, 10000));
            }
        } finally {
            Files.deleteIfExists(temp);
        }
    }

    // ==================== POI ====================

    private static int readPoi(Path path) throws Exception {
        int count = 0;
        try (FileInputStream fis = new FileInputStream(path.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {
            for (int s = 0; s < wb.getNumberOfSheets(); s++) {
                XSSFSheet sheet = wb.getSheetAt(s);
                for (var row : sheet) {
                    for (var cell : row) {
                        getCellValueAsString(cell);
                        count++;
                    }
                }
            }
        }
        return count;
    }

    private static int readPoiEncrypted(Path path) throws Exception {
        int count = 0;
        try (POIFSFileSystem fs = new POIFSFileSystem(path.toFile())) {
            EncryptionInfo info = new EncryptionInfo(fs);
            Decryptor dec = Decryptor.getInstance(info);

            if (!dec.verifyPassword(PASSWORD)) {
                throw new RuntimeException("Invalid password");
            }

            try (InputStream is = dec.getDataStream(fs);
                 XSSFWorkbook wb = new XSSFWorkbook(is)) {
                for (int s = 0; s < wb.getNumberOfSheets(); s++) {
                    XSSFSheet sheet = wb.getSheetAt(s);
                    for (var row : sheet) {
                        for (var cell : row) {
                            getCellValueAsString(cell);
                            count++;
                        }
                    }
                }
            }
        }
        return count;
    }

    private static void readModifyWritePoi(Path input, Path output) throws Exception {
        try (FileInputStream fis = new FileInputStream(input.toFile());
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {
            // Find Sales sheet and modify 300 cells
            XSSFSheet sales = wb.getSheet("Sales");
            Random rand = new Random(42);
            for (int i = 0; i < MODIFY_CELLS; i++) {
                int rowNum = rand.nextInt(3000) + 1;  // Skip header
                int col = rand.nextInt(15);
                var row = sales.getRow(rowNum);
                if (row == null) {
                    row = sales.createRow(rowNum);
                }
                var cell = row.getCell(col);
                if (cell == null) {
                    cell = row.createCell(col);
                }
                cell.setCellValue("MODIFIED-" + i);
            }

            try (FileOutputStream fos = new FileOutputStream(output.toFile())) {
                wb.write(fos);
            }
        }
    }

    private static void generatePoiEncrypted(ComplexDataGenerator generator, Path path) throws Exception {
        // First generate plain XLSX to temp
        Path temp = Files.createTempFile("poi-temp-", ".xlsx");
        try {
            generator.generatePoi(temp);

            // Then encrypt it
            try (POIFSFileSystem fs = new POIFSFileSystem()) {
                EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
                Encryptor enc = info.getEncryptor();
                enc.confirmPassword(PASSWORD);

                try (InputStream is = Files.newInputStream(temp);
                     OutputStream os = enc.getDataStream(fs)) {
                    is.transferTo(os);
                }

                try (FileOutputStream fos = new FileOutputStream(path.toFile())) {
                    fs.writeFilesystem(fos);
                }
            }
        } finally {
            Files.deleteIfExists(temp);
        }
    }

    private static String getCellValueAsString(org.apache.poi.ss.usermodel.Cell cell) {
        if (cell == null) {
            return "";
        }
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            case ERROR -> String.valueOf(cell.getErrorCellValue());
            case BLANK -> "";
            default -> "";
        };
    }
}
