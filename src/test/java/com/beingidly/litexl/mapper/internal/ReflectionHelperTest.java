package com.beingidly.litexl.mapper.internal;

import com.beingidly.litexl.*;
import com.beingidly.litexl.mapper.*;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class ReflectionHelperTest {

    @LitexlRow
    record Person(
        @LitexlColumn(index = 0, header = "Name") String name,
        @LitexlColumn(index = 1) int age
    ) {}

    @LitexlWorkbook
    record TestWorkbook(
        @LitexlSheet(name = "People") List<Person> people
    ) {}

    static class MutablePerson {
        @LitexlColumn(index = 0)
        private String name;
        @LitexlColumn(index = 1)
        private int age;

        public MutablePerson() {}

        public String getName() { return name; }
        public void setName(String name) { this.name = name; }
        public int getAge() { return age; }
        public void setAge(int age) { this.age = age; }
    }

    @Test
    void createRecordInstance() {
        Object[] args = {"Alice", 30};
        var person = ReflectionHelper.createInstance(Person.class, args);
        assertEquals("Alice", person.name());
        assertEquals(30, person.age());
    }

    @Test
    void createClassInstance() {
        var person = ReflectionHelper.createInstance(MutablePerson.class, new Object[0]);
        assertNotNull(person);
    }

    @Test
    void setFieldValue() {
        var person = new MutablePerson();
        var field = ReflectionHelper.getAnnotatedFields(MutablePerson.class, LitexlColumn.class).get(0);
        ReflectionHelper.setFieldValue(person, field, "Bob");
        assertEquals("Bob", person.getName());
    }

    @Test
    void getFieldValue() {
        var person = new Person("Alice", 30);
        var fields = ReflectionHelper.getAnnotatedFields(Person.class, LitexlColumn.class);
        assertEquals("Alice", ReflectionHelper.getFieldValue(person, fields.get(0)));
        assertEquals(30, ReflectionHelper.getFieldValue(person, fields.get(1)));
    }

    @Test
    void getAnnotatedFields() {
        var fields = ReflectionHelper.getAnnotatedFields(Person.class, LitexlColumn.class);
        assertEquals(2, fields.size());
        assertEquals("name", fields.get(0).getName());
        assertEquals("age", fields.get(1).getName());
    }

    @Test
    void isRecord() {
        assertTrue(ReflectionHelper.isRecord(Person.class));
        assertFalse(ReflectionHelper.isRecord(MutablePerson.class));
    }

    @Test
    void getGenericListType() {
        var field = ReflectionHelper.getAnnotatedFields(TestWorkbook.class, LitexlSheet.class).get(0);
        var listType = ReflectionHelper.getGenericListType(field);
        assertEquals(Person.class, listType);
    }

    // ========== Additional Tests for Edge Cases ==========

    // Test class with private fields (no public setters)
    static class PrivateFieldsOnly {
        @LitexlColumn(index = 0)
        private String secretField;

        public PrivateFieldsOnly() {}

        public String getSecretField() { return secretField; }
    }

    @Test
    void setFieldValue_privateField_shouldWork() {
        var obj = new PrivateFieldsOnly();
        var fields = ReflectionHelper.getAnnotatedFields(PrivateFieldsOnly.class, LitexlColumn.class);
        assertEquals(1, fields.size());

        ReflectionHelper.setFieldValue(obj, fields.get(0), "secret value");
        assertEquals("secret value", obj.getSecretField());
    }

    @Test
    void getFieldValue_privateField_shouldWork() {
        var obj = new PrivateFieldsOnly();
        var fields = ReflectionHelper.getAnnotatedFields(PrivateFieldsOnly.class, LitexlColumn.class);
        ReflectionHelper.setFieldValue(obj, fields.get(0), "hidden");

        assertEquals("hidden", ReflectionHelper.getFieldValue(obj, fields.get(0)));
    }

    // Test class without default constructor
    static class NoDefaultConstructor {
        @LitexlColumn(index = 0)
        private String value;

        public NoDefaultConstructor(String value) {
            this.value = value;
        }
    }

    @Test
    void createInstance_classWithoutDefaultConstructor_shouldThrow() {
        var exception = assertThrows(LitexlMapperException.class, () ->
            ReflectionHelper.createInstance(NoDefaultConstructor.class, new Object[0])
        );
        assertTrue(exception.getMessage().contains("Failed to create instance"));
    }

    // Test class with static fields
    static class ClassWithStaticFields {
        @LitexlColumn(index = 0)
        private static String staticField = "static";

        @LitexlColumn(index = 1)
        private String instanceField;

        public ClassWithStaticFields() {}
    }

    @Test
    void getAnnotatedFields_includesStaticFields() {
        var fields = ReflectionHelper.getAnnotatedFields(ClassWithStaticFields.class, LitexlColumn.class);
        // Both static and instance fields should be returned (they're both annotated)
        assertEquals(2, fields.size());
    }

    @Test
    void setFieldValue_staticField_shouldWork() throws NoSuchFieldException {
        var field = ClassWithStaticFields.class.getDeclaredField("staticField");
        // Static field can be set via setFieldValue
        ReflectionHelper.setFieldValue(null, field, "new static value");
        // Note: setting static field with null object works in Java
    }

    // Test isList method
    static class ListContainer {
        @LitexlColumn(index = 0)
        private List<String> stringList;

        @LitexlColumn(index = 1)
        private ArrayList<Integer> arrayList;

        @LitexlColumn(index = 2)
        private String notAList;

        public ListContainer() {}
    }

    @Test
    void isList_withListField_returnsTrue() throws NoSuchFieldException {
        var field = ListContainer.class.getDeclaredField("stringList");
        assertTrue(ReflectionHelper.isList(field));
    }

    @Test
    void isList_withArrayListField_returnsTrue() throws NoSuchFieldException {
        var field = ListContainer.class.getDeclaredField("arrayList");
        assertTrue(ReflectionHelper.isList(field));
    }

    @Test
    void isList_withNonListField_returnsFalse() throws NoSuchFieldException {
        var field = ListContainer.class.getDeclaredField("notAList");
        assertFalse(ReflectionHelper.isList(field));
    }

    // Test getGenericType method
    @Test
    void getGenericType_withParameterizedType_returnsGenericType() throws NoSuchFieldException {
        var field = ListContainer.class.getDeclaredField("stringList");
        var genericType = field.getGenericType();
        var result = ReflectionHelper.getGenericType(genericType);
        assertEquals(String.class, result);
    }

    @Test
    void getGenericType_withNonParameterizedType_throwsException() throws NoSuchFieldException {
        var field = ListContainer.class.getDeclaredField("notAList");
        var genericType = field.getGenericType();

        var exception = assertThrows(LitexlMapperException.class, () ->
            ReflectionHelper.getGenericType(genericType)
        );
        assertTrue(exception.getMessage().contains("Cannot determine generic type"));
    }

    // Test getGenericListType with non-parameterized field
    @Test
    void getGenericListType_withNonParameterizedField_throwsException() throws NoSuchFieldException {
        var field = ListContainer.class.getDeclaredField("notAList");

        var exception = assertThrows(LitexlMapperException.class, () ->
            ReflectionHelper.getGenericListType(field)
        );
        assertTrue(exception.getMessage().contains("Cannot determine generic type"));
    }

    // Test isEmptyRow method
    @Test
    void isEmptyRow_withAllEmptyCells_returnsTrue() {
        try (var wb = Workbook.create()) {
            var sheet = wb.addSheet("Test");
            var row = sheet.row(0);
            row.cell(0).setEmpty();
            row.cell(1).setEmpty();
            assertTrue(ReflectionHelper.isEmptyRow(row));
        }
    }

    @Test
    void isEmptyRow_withNonemptyCell_returnsFalse() {
        try (var wb = Workbook.create()) {
            var sheet = wb.addSheet("Test");
            var row = sheet.row(0);
            row.cell(0).set("value");
            row.cell(1).setEmpty();
            assertFalse(ReflectionHelper.isEmptyRow(row));
        }
    }

    @Test
    void isEmptyRow_withNumberCell_returnsFalse() {
        try (var wb = Workbook.create()) {
            var sheet = wb.addSheet("Test");
            var row = sheet.row(0);
            row.cell(0).set(42.0);
            assertFalse(ReflectionHelper.isEmptyRow(row));
        }
    }

    @Test
    void isEmptyRow_withEmptyRow_returnsTrue() {
        try (var wb = Workbook.create()) {
            var sheet = wb.addSheet("Test");
            var row = sheet.row(0);
            assertTrue(ReflectionHelper.isEmptyRow(row));
        }
    }

    // Test getAnnotatedFields with class having no annotated fields
    static class NoAnnotations {
        private String field1;
        private int field2;

        public NoAnnotations() {}
    }

    @Test
    void getAnnotatedFields_withNoAnnotatedFields_returnsEmptyList() {
        var fields = ReflectionHelper.getAnnotatedFields(NoAnnotations.class, LitexlColumn.class);
        assertTrue(fields.isEmpty());
    }

    // Test setFieldValue and getFieldValue with null value
    @Test
    void setFieldValue_withNullValue_setsNull() {
        var person = new MutablePerson();
        person.setName("Initial");
        var field = ReflectionHelper.getAnnotatedFields(MutablePerson.class, LitexlColumn.class).get(0);

        ReflectionHelper.setFieldValue(person, field, null);
        assertNull(person.getName());
    }

    @Test
    void getFieldValue_withNullValue_returnsNull() {
        var person = new MutablePerson();
        person.setName(null);
        var field = ReflectionHelper.getAnnotatedFields(MutablePerson.class, LitexlColumn.class).get(0);

        assertNull(ReflectionHelper.getFieldValue(person, field));
    }

    // Test createInstance with record having null arguments
    @Test
    void createRecordInstance_withNullArguments() {
        Object[] args = {null, 25};
        var person = ReflectionHelper.createInstance(Person.class, args);
        assertNull(person.name());
        assertEquals(25, person.age());
    }

    // Test getAnnotatedFields with LitexlSheet annotation
    @Test
    void getAnnotatedFields_forWorkbook_returnsSheetFields() {
        var fields = ReflectionHelper.getAnnotatedFields(TestWorkbook.class, LitexlSheet.class);
        assertEquals(1, fields.size());
        assertEquals("people", fields.get(0).getName());
    }

    // Test raw type List field
    static class RawListContainer {
        @SuppressWarnings("rawtypes")
        @LitexlColumn(index = 0)
        private List rawList;

        public RawListContainer() {}
    }

    @Test
    void getGenericListType_withRawType_throwsException() throws NoSuchFieldException {
        var field = RawListContainer.class.getDeclaredField("rawList");

        var exception = assertThrows(LitexlMapperException.class, () ->
            ReflectionHelper.getGenericListType(field)
        );
        assertTrue(exception.getMessage().contains("Cannot determine generic type"));
    }
}
