package com.beingidly.litexl.mapper;

import org.junit.jupiter.api.Test;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import static org.junit.jupiter.api.Assertions.*;

class AnnotationTest {

    @LitexlWorkbook
    record TestWorkbook(
        @LitexlSheet(name = "Sheet1")
        java.util.List<TestRow> rows
    ) {}

    @LitexlRow
    record TestRow(
        @LitexlColumn(index = 0, header = "Name")
        String name,
        @LitexlColumn(index = 1)
        int age
    ) {}

    @LitexlSheet
    record TestSummary(
        @LitexlCell(row = 0, column = 0)
        String title
    ) {}

    @Test
    void workbookAnnotationPresent() {
        assertTrue(TestWorkbook.class.isAnnotationPresent(LitexlWorkbook.class));
    }

    @Test
    void sheetAnnotationOnField() throws NoSuchFieldException {
        var field = TestWorkbook.class.getDeclaredField("rows");
        var ann = field.getAnnotation(LitexlSheet.class);
        assertNotNull(ann);
        assertEquals("Sheet1", ann.name());
        assertEquals(-1, ann.index());
        assertEquals(0, ann.headerRow());
        assertEquals(1, ann.dataStartRow());
        assertEquals(0, ann.dataStartColumn());
        assertEquals(RegionDetection.NONE, ann.regionDetection());
    }

    @Test
    void rowAnnotationPresent() {
        assertTrue(TestRow.class.isAnnotationPresent(LitexlRow.class));
    }

    @Test
    void columnAnnotationOnField() throws NoSuchFieldException {
        var nameField = TestRow.class.getDeclaredField("name");
        var ann = nameField.getAnnotation(LitexlColumn.class);
        assertNotNull(ann);
        assertEquals(0, ann.index());
        assertEquals("Name", ann.header());
    }

    @Test
    void cellAnnotationOnField() throws NoSuchFieldException {
        var field = TestSummary.class.getDeclaredField("title");
        var ann = field.getAnnotation(LitexlCell.class);
        assertNotNull(ann);
        assertEquals(0, ann.row());
        assertEquals(0, ann.column());
    }

    @Test
    void annotationsRetainedAtRuntime() {
        assertEquals(RetentionPolicy.RUNTIME,
            LitexlWorkbook.class.getAnnotation(Retention.class).value());
        assertEquals(RetentionPolicy.RUNTIME,
            LitexlSheet.class.getAnnotation(Retention.class).value());
        assertEquals(RetentionPolicy.RUNTIME,
            LitexlRow.class.getAnnotation(Retention.class).value());
        assertEquals(RetentionPolicy.RUNTIME,
            LitexlColumn.class.getAnnotation(Retention.class).value());
        assertEquals(RetentionPolicy.RUNTIME,
            LitexlCell.class.getAnnotation(Retention.class).value());
    }
}
