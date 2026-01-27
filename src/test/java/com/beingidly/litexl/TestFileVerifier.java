package com.beingidly.litexl;

import java.nio.file.Path;

/**
 * Verifies that the generated test files can be read correctly.
 */
public class TestFileVerifier {

    public static void main(String[] args) {
        try {
            // Test unencrypted file
            System.out.println("=== Testing unencrypted file ===");
            try (Workbook wb = Workbook.open(Path.of("src/test/resources/test-data.xlsx"))) {
                Sheet sheet = wb.getSheet(0);
                System.out.println("Sheet: " + sheet.name());
                System.out.println("Cell A1: " + sheet.cell(0, 0).string());
                System.out.println("Cell B1: " + sheet.cell(0, 1).number());
                System.out.println("Cell A2: " + sheet.cell(1, 0).bool());
            }

            // Test AES-128 encrypted file
            System.out.println("\n=== Testing AES-128 encrypted file ===");
            try (Workbook wb = Workbook.open(Path.of("src/test/resources/encrypted/aes128.xlsx"), "password123")) {
                Sheet sheet = wb.getSheet(0);
                System.out.println("Sheet: " + sheet.name());
                System.out.println("Cell A1: " + sheet.cell(0, 0).string());
                System.out.println("Cell B1: " + sheet.cell(0, 1).number());
                System.out.println("Cell A2: " + sheet.cell(1, 0).bool());
            }

            // Test AES-256 encrypted file
            System.out.println("\n=== Testing AES-256 encrypted file ===");
            try (Workbook wb = Workbook.open(Path.of("src/test/resources/encrypted/aes256.xlsx"), "password123")) {
                Sheet sheet = wb.getSheet(0);
                System.out.println("Sheet: " + sheet.name());
                System.out.println("Cell A1: " + sheet.cell(0, 0).string());
                System.out.println("Cell B1: " + sheet.cell(0, 1).number());
                System.out.println("Cell A2: " + sheet.cell(1, 0).bool());
            }

            System.out.println("\n=== All files verified successfully! ===");

        } catch (Exception e) {
            System.err.println("VERIFICATION FAILED: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
        }
    }
}
