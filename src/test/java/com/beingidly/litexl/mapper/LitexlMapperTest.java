package com.beingidly.litexl.mapper;

import com.beingidly.litexl.*;
import com.beingidly.litexl.crypto.EncryptionOptions;
import com.beingidly.litexl.style.Style;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;
import java.time.LocalDateTime;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class LitexlMapperTest {

    @TempDir
    Path tempDir;

    @LitexlRow
    record Person(
        @LitexlColumn(index = 0, header = "Name") String name,
        @LitexlColumn(index = 1, header = "Age") int age
    ) {}

    @LitexlWorkbook
    record SimpleWorkbook(
        @LitexlSheet(name = "People")
        List<Person> people
    ) {}

    @Test
    void writeAndRead() throws Exception {
        var data = new SimpleWorkbook(List.of(
            new Person("Alice", 30),
            new Person("Bob", 25)
        ));

        var path = tempDir.resolve("test.xlsx");
        LitexlMapper.write(data, path);
        assertTrue(path.toFile().exists());

        var result = LitexlMapper.read(path, SimpleWorkbook.class);
        assertEquals(2, result.people().size());
        assertEquals("Alice", result.people().get(0).name());
        assertEquals(30, result.people().get(0).age());
    }

    @Test
    void builderPattern() throws Exception {
        var data = new SimpleWorkbook(List.of(new Person("Test", 20)));
        var path = tempDir.resolve("builder.xlsx");

        var mapper = LitexlMapper.builder()
            .dateFormat("yyyy/MM/dd")
            .nullStrategy(NullStrategy.EMPTY_CELL)
            .build();

        mapper.writeFile(data, path);
        var result = mapper.readFile(path, SimpleWorkbook.class);

        assertEquals(1, result.people().size());
    }

    @Test
    void staticReadMethod() throws Exception {
        var path = tempDir.resolve("static.xlsx");

        try (Workbook wb = Workbook.create()) {
            var sheet = wb.addSheet("People");
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Age");
            sheet.cell(1, 0).set("Charlie");
            sheet.cell(1, 1).set(35.0);
            wb.save(path);
        }

        var result = LitexlMapper.read(path, SimpleWorkbook.class);
        assertEquals(1, result.people().size());
        assertEquals("Charlie", result.people().get(0).name());
    }

    @Test
    void readEncryptedFile() throws Exception {
        var path = tempDir.resolve("encrypted.xlsx");
        var password = "testPassword123";

        // Create an encrypted file using low-level API
        try (Workbook wb = Workbook.create()) {
            var sheet = wb.addSheet("People");
            sheet.cell(0, 0).set("Name");
            sheet.cell(0, 1).set("Age");
            sheet.cell(1, 0).set("EncryptedUser");
            sheet.cell(1, 1).set(42.0);
            wb.save(path, EncryptionOptions.aes256(password));
        }

        // Read encrypted file using LitexlMapper with password
        var result = LitexlMapper.read(path, password, SimpleWorkbook.class);
        assertEquals(1, result.people().size());
        assertEquals("EncryptedUser", result.people().get(0).name());
        assertEquals(42, result.people().get(0).age());
    }

    @Test
    void writeEncryptedFile() throws Exception {
        var data = new SimpleWorkbook(List.of(
            new Person("SecureAlice", 28),
            new Person("SecureBob", 33)
        ));

        var path = tempDir.resolve("mapper_encrypted.xlsx");
        var password = "securePassword456";
        var encryptionOptions = EncryptionOptions.aes256(password);

        // Write with encryption
        LitexlMapper.write(data, path, encryptionOptions);
        assertTrue(path.toFile().exists());

        // Read back with password to verify
        var result = LitexlMapper.read(path, password, SimpleWorkbook.class);
        assertEquals(2, result.people().size());
        assertEquals("SecureAlice", result.people().get(0).name());
        assertEquals(28, result.people().get(0).age());
        assertEquals("SecureBob", result.people().get(1).name());
        assertEquals(33, result.people().get(1).age());
    }

    @Test
    void builderMethodChaining() {
        // Test that builder methods return the same builder instance for chaining
        var builder1 = LitexlMapper.builder();
        var builder2 = builder1.dateFormat("yyyy/MM/dd");
        var builder3 = builder2.nullStrategy(NullStrategy.EMPTY_CELL);

        // All should be the same builder instance
        assertSame(builder1, builder2);
        assertSame(builder2, builder3);

        // Build should return a non-null mapper
        var mapper = builder3.build();
        assertNotNull(mapper);
    }

    @Test
    void registerClassForGraalVM() {
        // Record initial state
        var initialClasses = LitexlMapper.getRegisteredClasses();

        // Register classes
        LitexlMapper.register(Person.class, SimpleWorkbook.class);

        // Verify registration
        var registeredClasses = LitexlMapper.getRegisteredClasses();
        assertTrue(registeredClasses.contains(Person.class));
        assertTrue(registeredClasses.contains(SimpleWorkbook.class));

        // getRegisteredClasses should return an immutable copy
        assertThrows(UnsupportedOperationException.class, () -> {
            registeredClasses.add(String.class);
        });
    }
}
