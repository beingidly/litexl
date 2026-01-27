package com.beingidly.litexl.examples.mapper;

import com.beingidly.litexl.examples.util.ExampleUtils;
import com.beingidly.litexl.mapper.LitexlColumn;
import com.beingidly.litexl.mapper.LitexlMapper;
import com.beingidly.litexl.mapper.LitexlRow;
import com.beingidly.litexl.mapper.LitexlSheet;
import com.beingidly.litexl.mapper.LitexlWorkbook;

import java.nio.file.Path;
import java.util.List;

/**
 * Example: Simple object-to-Excel mapping.
 *
 * This example demonstrates:
 * - Defining workbook and row records
 * - Mapping fields to columns
 * - Writing objects to Excel
 * - Reading Excel back to objects
 */
public class Ex01_SimpleMapping {

    // Define a row record
    @LitexlRow
    public record Person(
        @LitexlColumn(index = 0, header = "Name")
        String name,

        @LitexlColumn(index = 1, header = "Age")
        int age,

        @LitexlColumn(index = 2, header = "Email")
        String email
    ) {}

    // Define a workbook record
    @LitexlWorkbook
    public record PeopleReport(
        @LitexlSheet(name = "People")
        List<Person> people
    ) {}

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex01_simple_mapping.xlsx");

        // Create sample data
        List<Person> people = List.of(
            new Person("Alice", 30, "alice@example.com"),
            new Person("Bob", 25, "bob@example.com"),
            new Person("Charlie", 35, "charlie@example.com")
        );

        PeopleReport report = new PeopleReport(people);

        // Write to Excel
        LitexlMapper.write(report, outputPath);
        ExampleUtils.printCreated(outputPath);

        // Read back from Excel
        System.out.println();
        System.out.println("Reading back from Excel:");

        PeopleReport readBack = LitexlMapper.read(outputPath, PeopleReport.class);

        for (Person person : readBack.people()) {
            System.out.println("  " + person.name() + ", " + person.age() + ", " + person.email());
        }
    }
}
