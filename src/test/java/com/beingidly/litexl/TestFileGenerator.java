package com.beingidly.litexl;

import com.beingidly.litexl.crypto.EncryptionOptions;

import java.nio.file.Files;
import java.nio.file.Path;

/**
 * Utility to generate test files for encrypted file testing.
 * Run this to create test resources in src/test/resources/.
 */
public class TestFileGenerator {

    public static void main(String[] args) throws Exception {
        Path baseDir = Path.of("src/test/resources");
        Path encryptedDir = baseDir.resolve("encrypted");

        // Create directories if needed
        Files.createDirectories(encryptedDir);

        // Create workbook with test data
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("TestSheet");

            // Row 0: String and Number
            sheet.cell(0, 0).set("Hello");
            sheet.cell(0, 1).set(123.45);

            // Row 1: Boolean
            sheet.cell(1, 0).set(true);

            // Save unencrypted for baseline
            workbook.save(baseDir.resolve("test-data.xlsx"));
            System.out.println("Created: " + baseDir.resolve("test-data.xlsx"));

            // Save AES-128 encrypted
            workbook.save(
                encryptedDir.resolve("aes128.xlsx"),
                new EncryptionOptions(EncryptionOptions.Algorithm.AES_128, "password123", 100000)
            );
            System.out.println("Created: " + encryptedDir.resolve("aes128.xlsx"));

            // Save AES-256 encrypted
            workbook.save(
                encryptedDir.resolve("aes256.xlsx"),
                new EncryptionOptions(EncryptionOptions.Algorithm.AES_256, "password123", 100000)
            );
            System.out.println("Created: " + encryptedDir.resolve("aes256.xlsx"));
        }

        System.out.println("\nTest files generated successfully!");
    }
}
