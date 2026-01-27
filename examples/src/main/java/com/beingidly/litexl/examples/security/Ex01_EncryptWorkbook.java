package com.beingidly.litexl.examples.security;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.crypto.EncryptionOptions;
import com.beingidly.litexl.examples.util.ExampleUtils;

import java.nio.file.Path;

/**
 * Example: Encrypting workbooks with AES.
 *
 * This example demonstrates:
 * - Saving a workbook with AES-256 encryption
 * - Opening an encrypted workbook with password
 */
public class Ex01_EncryptWorkbook {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex01_encrypted.xlsx");
        String password = "secret123";

        // Create and encrypt a workbook
        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Confidential");

            sheet.cell(0, 0).set("Confidential Data");
            sheet.cell(1, 0).set("Account Number");
            sheet.cell(1, 1).set("1234-5678-9012");
            sheet.cell(2, 0).set("Balance");
            sheet.cell(2, 1).set(50000.00);

            // Save with AES-256 encryption
            workbook.save(outputPath, EncryptionOptions.aes256(password));
        }

        ExampleUtils.printCreated(outputPath);
        System.out.println("Password: " + password);
        System.out.println();

        // Open the encrypted workbook
        System.out.println("Opening encrypted workbook with password...");

        try (Workbook workbook = Workbook.open(outputPath, password)) {
            Sheet sheet = workbook.getSheet(0);
            System.out.println("  Title: " + sheet.cell(0, 0).string());
            System.out.println("  " + sheet.cell(1, 0).string() + ": " + sheet.cell(1, 1).string());
            System.out.println("  " + sheet.cell(2, 0).string() + ": " + sheet.cell(2, 1).number());
        }

        System.out.println();
        System.out.println("Try opening the file in Excel - it will ask for the password.");
    }
}
