package com.beingidly.litexl.mapper.internal;

import com.beingidly.litexl.*;
import com.beingidly.litexl.mapper.*;
import org.jspecify.annotations.Nullable;

import java.lang.reflect.Field;
import java.lang.reflect.RecordComponent;
import java.util.Comparator;
import java.util.List;

/**
 * Writes Java objects to an Excel workbook using annotations.
 */
public final class WorkbookWriter {

    @SuppressWarnings("unused")
    public WorkbookWriter(MapperConfig config) {
        // config reserved for future extensibility
    }

    /**
     * Writes a Java object to the workbook.
     *
     * @param workbook the target workbook
     * @param data the data object (must be annotated with @LitexlWorkbook)
     * @throws LitexlMapperException if writing fails
     */
    public void write(Workbook workbook, Object data) {
        Class<?> type = data.getClass();
        validateWorkbookAnnotation(type);

        if (ReflectionHelper.isRecord(type)) {
            writeRecord(workbook, data, type);
        } else {
            writeClass(workbook, data, type);
        }
    }

    private void validateWorkbookAnnotation(Class<?> type) {
        if (!type.isAnnotationPresent(LitexlWorkbook.class)) {
            throw new LitexlMapperException("Class " + type.getName() + " is not annotated with @LitexlWorkbook");
        }
    }

    private void validateRowAnnotation(Class<?> type) {
        if (!type.isAnnotationPresent(LitexlRow.class)) {
            throw new LitexlMapperException("Class " + type.getName() + " is not annotated with @LitexlRow");
        }
    }

    private void writeRecord(Workbook workbook, Object data, Class<?> type) {
        RecordComponent[] components = type.getRecordComponents();

        for (RecordComponent component : components) {
            if (component.isAnnotationPresent(LitexlSheet.class)) {
                LitexlSheet sheetAnnotation = component.getAnnotation(LitexlSheet.class);
                Object fieldValue = getRecordComponentValue(data, component);
                writeSheetData(workbook, sheetAnnotation, component.getName(), fieldValue, component.getType(), component.getGenericType());
            } else if (component.isAnnotationPresent(LitexlCell.class)) {
                LitexlCell cellAnnotation = component.getAnnotation(LitexlCell.class);
                Object fieldValue = getRecordComponentValue(data, component);
                writeCellData(workbook, cellAnnotation, fieldValue, component.getType());
            }
        }
    }

    private void writeClass(Workbook workbook, Object data, Class<?> type) {
        // Process @LitexlSheet annotated fields
        List<Field> sheetFields = ReflectionHelper.getAnnotatedFields(type, LitexlSheet.class);
        for (Field field : sheetFields) {
            LitexlSheet sheetAnnotation = field.getAnnotation(LitexlSheet.class);
            Object fieldValue = ReflectionHelper.getFieldValue(data, field);
            writeSheetData(workbook, sheetAnnotation, field.getName(), fieldValue, field.getType(), field.getGenericType());
        }

        // Process @LitexlCell annotated fields
        List<Field> cellFields = ReflectionHelper.getAnnotatedFields(type, LitexlCell.class);
        cellFields.sort(Comparator
            .comparingInt((Field f) -> f.getAnnotation(LitexlCell.class).row())
            .thenComparingInt(f -> f.getAnnotation(LitexlCell.class).column()));
        for (Field field : cellFields) {
            LitexlCell cellAnnotation = field.getAnnotation(LitexlCell.class);
            Object fieldValue = ReflectionHelper.getFieldValue(data, field);
            writeCellData(workbook, cellAnnotation, fieldValue, field.getType());
        }
    }

    private @Nullable Object getRecordComponentValue(Object data, RecordComponent component) {
        try {
            var accessor = component.getAccessor();
            accessor.setAccessible(true);
            return accessor.invoke(data);
        } catch (ReflectiveOperationException e) {
            throw new LitexlMapperException("Failed to get record component value: " + component.getName(), e);
        }
    }

    private void writeSheetData(Workbook workbook, LitexlSheet sheetAnnotation, String fieldName,
                                @Nullable Object fieldValue, Class<?> fieldType, java.lang.reflect.Type genericType) {
        // Determine sheet name
        String sheetName = sheetAnnotation.name().isEmpty() ? fieldName : sheetAnnotation.name();

        // Create sheet
        Sheet sheet = workbook.addSheet(sheetName);

        if (fieldValue == null) {
            return;
        }

        if (List.class.isAssignableFrom(fieldType)) {
            Class<?> elementType = ReflectionHelper.getGenericType(genericType);
            writeRowList(sheet, sheetAnnotation, (List<?>) fieldValue, elementType);
        } else {
            // Single object - write to a single row
            writeSingleObject(sheet, sheetAnnotation, fieldValue);
        }
    }

    private void writeCellData(Workbook workbook, LitexlCell cellAnnotation,
                               @Nullable Object fieldValue, Class<?> fieldType) {
        // For @LitexlCell at workbook level, we need a sheet - use first or create one
        Sheet sheet;
        if (workbook.sheetCount() == 0) {
            sheet = workbook.addSheet("Sheet1");
        } else {
            sheet = workbook.getSheet(0);
        }

        if (sheet == null) {
            return;
        }

        CellValue cellValue = convertToCell(fieldValue, fieldType, cellAnnotation.converter());
        Cell cell = sheet.cell(cellAnnotation.row(), cellAnnotation.column());
        cell.setValue(cellValue);
    }

    private void writeRowList(Sheet sheet, LitexlSheet sheetAnnotation, List<?> items, Class<?> elementType) {
        validateRowAnnotation(elementType);
        int headerRow = sheetAnnotation.headerRow();
        int dataStartRow = sheetAnnotation.dataStartRow();
        int dataStartColumn = sheetAnnotation.dataStartColumn();

        // Write header row
        writeHeaderRow(sheet, headerRow, dataStartColumn, elementType);

        // Write data rows
        int currentRow = dataStartRow;
        for (Object item : items) {
            writeDataRow(sheet, currentRow, dataStartColumn, item, elementType);
            currentRow++;
        }
    }

    private void writeHeaderRow(Sheet sheet, int rowNum, int startColumn, Class<?> elementType) {
        if (ReflectionHelper.isRecord(elementType)) {
            RecordComponent[] components = elementType.getRecordComponents();
            for (RecordComponent component : components) {
                LitexlColumn columnAnnotation = component.getAnnotation(LitexlColumn.class);
                if (columnAnnotation != null) {
                    int colIndex = columnAnnotation.index();
                    String header = columnAnnotation.header().isEmpty() ? component.getName() : columnAnnotation.header();
                    sheet.cell(rowNum, startColumn + colIndex).set(header);
                }
            }
        } else {
            List<Field> columnFields = ReflectionHelper.getAnnotatedFields(elementType, LitexlColumn.class);
            for (Field field : columnFields) {
                LitexlColumn columnAnnotation = field.getAnnotation(LitexlColumn.class);
                int colIndex = columnAnnotation.index();
                String header = columnAnnotation.header().isEmpty() ? field.getName() : columnAnnotation.header();
                sheet.cell(rowNum, startColumn + colIndex).set(header);
            }
        }
    }

    private void writeDataRow(Sheet sheet, int rowNum, int startColumn, Object item, Class<?> elementType) {
        if (ReflectionHelper.isRecord(elementType)) {
            writeRecordDataRow(sheet, rowNum, startColumn, item);
        } else {
            writeClassDataRow(sheet, rowNum, startColumn, item);
        }
    }

    private void writeRecordDataRow(Sheet sheet, int rowNum, int startColumn, Object item) {
        Class<?> type = item.getClass();
        RecordComponent[] components = type.getRecordComponents();

        for (RecordComponent component : components) {
            LitexlColumn columnAnnotation = component.getAnnotation(LitexlColumn.class);
            if (columnAnnotation != null) {
                int colIndex = columnAnnotation.index();
                Object value = getRecordComponentValue(item, component);
                CellValue cellValue = convertToCell(value, component.getType(), columnAnnotation.converter());
                setCellValue(sheet.cell(rowNum, startColumn + colIndex), cellValue);
            }
        }
    }

    private void writeClassDataRow(Sheet sheet, int rowNum, int startColumn, Object item) {
        List<Field> columnFields = ReflectionHelper.getAnnotatedFields(item.getClass(), LitexlColumn.class);

        for (Field field : columnFields) {
            LitexlColumn columnAnnotation = field.getAnnotation(LitexlColumn.class);
            int colIndex = columnAnnotation.index();
            Object value = ReflectionHelper.getFieldValue(item, field);
            CellValue cellValue = convertToCell(value, field.getType(), columnAnnotation.converter());
            setCellValue(sheet.cell(rowNum, startColumn + colIndex), cellValue);
        }
    }

    private void writeSingleObject(Sheet sheet, LitexlSheet sheetAnnotation, Object item) {
        validateRowAnnotation(item.getClass());
        int dataStartRow = sheetAnnotation.dataStartRow();
        int dataStartColumn = sheetAnnotation.dataStartColumn();
        Class<?> type = item.getClass();

        // Write header if applicable
        writeHeaderRow(sheet, sheetAnnotation.headerRow(), dataStartColumn, type);

        // Write data
        writeDataRow(sheet, dataStartRow, dataStartColumn, item, type);
    }

    @SuppressWarnings("unchecked")
    private CellValue convertToCell(@Nullable Object value, Class<?> type,
                                    Class<? extends LitexlConverter<?>> converterClass) {
        // Check for custom converter
        if (converterClass != LitexlConverter.None.class) {
            try {
                LitexlConverter<Object> converter = (LitexlConverter<Object>) converterClass.getDeclaredConstructor().newInstance();
                return converter.toCell(value);
            } catch (ReflectiveOperationException e) {
                throw new LitexlMapperException("Failed to create converter: " + converterClass.getName(), e);
            }
        }

        if (value == null) {
            return new CellValue.Empty();
        }

        // Use TypeConverter for basic types
        if (TypeConverter.isSupported(type)) {
            return TypeConverter.toCell(value, type);
        }

        throw new LitexlMapperException("Unsupported type: " + type.getName() + ". Use a custom converter.");
    }

    private void setCellValue(Cell cell, CellValue cellValue) {
        cell.setValue(cellValue);
    }
}
