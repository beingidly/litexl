package com.beingidly.litexl.mapper.internal;

import com.beingidly.litexl.*;
import com.beingidly.litexl.mapper.*;
import org.jspecify.annotations.Nullable;

import java.lang.reflect.Field;
import java.lang.reflect.RecordComponent;
import java.util.*;

import static java.util.Collections.emptySet;

/**
 * Reads data from an Excel workbook and maps it to Java objects using annotations.
 */
public final class WorkbookReader {

    private final MapperConfig config;

    public WorkbookReader(MapperConfig config) {
        this.config = config;
    }

    /**
     * Reads a workbook and maps it to the specified type.
     *
     * @param workbook the source workbook
     * @param type the target type (must be annotated with @LitexlWorkbook)
     * @return an instance of the target type populated with data from the workbook
     * @throws LitexlMapperException if mapping fails
     */
    public <T> T read(Workbook workbook, Class<T> type) {
        validateWorkbookAnnotation(type);

        if (ReflectionHelper.isRecord(type)) {
            return readRecord(workbook, type);
        } else {
            return readClass(workbook, type);
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

    private <T> T readRecord(Workbook workbook, Class<T> type) {
        RecordComponent[] components = type.getRecordComponents();
        Object[] args = new Object[components.length];

        for (int i = 0; i < components.length; i++) {
            RecordComponent component = components[i];
            args[i] = readComponent(workbook, component);
        }

        return ReflectionHelper.createInstance(type, args);
    }

    private <T> T readClass(Workbook workbook, Class<T> type) {
        T instance = ReflectionHelper.createInstance(type, null);

        // Process @LitexlSheet annotated fields
        List<Field> sheetFields = ReflectionHelper.getAnnotatedFields(type, LitexlSheet.class);
        for (Field field : sheetFields) {
            Object value = readSheetField(workbook, field);
            ReflectionHelper.setFieldValue(instance, field, value);
        }

        // Process @LitexlCell annotated fields
        List<Field> cellFields = ReflectionHelper.getAnnotatedFields(type, LitexlCell.class);
        for (Field field : cellFields) {
            Object value = readCellField(workbook, field);
            ReflectionHelper.setFieldValue(instance, field, value);
        }

        return instance;
    }

    private @Nullable Object readComponent(Workbook workbook, RecordComponent component) {
        if (component.isAnnotationPresent(LitexlSheet.class)) {
            LitexlSheet sheetAnnotation = component.getAnnotation(LitexlSheet.class);
            return readSheetData(workbook, sheetAnnotation, component.getType(), component.getGenericType());
        } else if (component.isAnnotationPresent(LitexlCell.class)) {
            LitexlCell cellAnnotation = component.getAnnotation(LitexlCell.class);
            return readCellData(workbook, cellAnnotation, component.getType());
        }
        return null;
    }

    private @Nullable Object readSheetField(Workbook workbook, Field field) {
        LitexlSheet sheetAnnotation = field.getAnnotation(LitexlSheet.class);
        return readSheetData(workbook, sheetAnnotation, field.getType(), field.getGenericType());
    }

    private @Nullable Object readCellField(Workbook workbook, Field field) {
        LitexlCell cellAnnotation = field.getAnnotation(LitexlCell.class);
        return readCellData(workbook, cellAnnotation, field.getType());
    }

    private @Nullable Object readSheetData(Workbook workbook, LitexlSheet sheetAnnotation,
                                           Class<?> fieldType, java.lang.reflect.Type genericType) {
        Sheet sheet = findSheet(workbook, sheetAnnotation);
        if (sheet == null) {
            return null;
        }

        if (List.class.isAssignableFrom(fieldType)) {
            Class<?> elementType = ReflectionHelper.getGenericType(genericType);
            return readRowList(sheet, sheetAnnotation, elementType);
        } else {
            // Single object - read from a single row or use @LitexlCell
            return readSingleObject(sheet, sheetAnnotation, fieldType);
        }
    }

    private @Nullable Object readCellData(Workbook workbook, LitexlCell cellAnnotation, Class<?> fieldType) {
        // For @LitexlCell at workbook level, we need the first sheet or a specified sheet
        Sheet sheet = workbook.getSheet(0);
        if (sheet == null) {
            return null;
        }

        Cell cell = sheet.getCell(cellAnnotation.row(), cellAnnotation.column());
        if (cell == null) {
            return null;
        }

        return convertCellValue(cell.value(), fieldType, cellAnnotation.converter());
    }

    private @Nullable Sheet findSheet(Workbook workbook, LitexlSheet sheetAnnotation) {
        if (!sheetAnnotation.name().isEmpty()) {
            return workbook.getSheet(sheetAnnotation.name());
        } else if (sheetAnnotation.index() >= 0) {
            return workbook.getSheet(sheetAnnotation.index());
        }
        // Default to first sheet
        return workbook.getSheet(0);
    }

    private List<Object> readRowList(Sheet sheet, LitexlSheet sheetAnnotation, Class<?> elementType) {
        validateRowAnnotation(elementType);

        // AUTO detection: find header row dynamically
        if (sheetAnnotation.regionDetection() == RegionDetection.AUTO) {
            Set<String> headers = extractHeadersFromType(elementType);
            if (!headers.isEmpty()) {
                RegionDetector.Region region = RegionDetector.detectRegion(sheet, headers, 0);
                if (region != null) {
                    return readRowsInRegion(sheet, elementType, region);
                }
                return List.of();  // headers not found
            }
        }

        // NONE mode: use fixed dataStartRow
        List<Object> result = new ArrayList<>();
        int dataStartRow = sheetAnnotation.dataStartRow();

        sheet.forEachRow(row -> {
            int rowNum = row.rowNum();
            if (rowNum < dataStartRow) {
                return true;
            }

            Object rowObject = readRow(row, elementType);
            if (rowObject != null) {
                result.add(rowObject);
            }
            return true;
        });

        return result;
    }

    private @Nullable Object readRow(Row row, Class<?> type) {
        if (ReflectionHelper.isRecord(type)) {
            return readRowAsRecord(row, type);
        } else {
            return readRowAsClass(row, type);
        }
    }

    private @Nullable Object readRowAsRecord(Row row, Class<?> type) {
        RecordComponent[] components = type.getRecordComponents();
        Object[] args = new Object[components.length];

        for (int i = 0; i < components.length; i++) {
            RecordComponent component = components[i];
            LitexlColumn columnAnnotation = component.getAnnotation(LitexlColumn.class);

            if (columnAnnotation != null) {
                int colIndex = columnAnnotation.index();
                Cell cell = row.getCell(colIndex);
                CellValue cellValue = cell != null ? cell.value() : new CellValue.Empty();
                args[i] = convertCellValue(cellValue, component.getType(), columnAnnotation.converter());
            } else {
                args[i] = getDefaultValue(component.getType());
            }
        }

        return ReflectionHelper.createInstance(type, args);
    }

    private @Nullable Object readRowAsClass(Row row, Class<?> type) {
        Object instance = ReflectionHelper.createInstance(type, null);

        List<Field> columnFields = ReflectionHelper.getAnnotatedFields(type, LitexlColumn.class);
        for (Field field : columnFields) {
            LitexlColumn columnAnnotation = field.getAnnotation(LitexlColumn.class);
            int colIndex = columnAnnotation.index();

            Cell cell = row.getCell(colIndex);
            CellValue cellValue = cell != null ? cell.value() : new CellValue.Empty();
            Object value = convertCellValue(cellValue, field.getType(), columnAnnotation.converter());
            ReflectionHelper.setFieldValue(instance, field, value);
        }

        return instance;
    }

    private @Nullable Object readSingleObject(Sheet sheet, LitexlSheet sheetAnnotation, Class<?> type) {
        // Check if this is a multi-region sheet
        if (sheetAnnotation.regionDetection() == RegionDetection.AUTO) {
            // Find all List fields in the type
            List<Field> listFields = getListFields(type);
            if (!listFields.isEmpty()) {
                return readMultiRegionSheet(sheet, type, listFields);
            }
        }

        // Fall back to single row reading
        validateRowAnnotation(type);
        int dataStartRow = sheetAnnotation.dataStartRow();
        Row row = sheet.getRow(dataStartRow);
        if (row == null) {
            return null;
        }
        return readRow(row, type);
    }

    private List<Field> getListFields(Class<?> type) {
        List<Field> result = new ArrayList<>();
        if (type.isRecord()) {
            for (RecordComponent component : type.getRecordComponents()) {
                if (List.class.isAssignableFrom(component.getType())) {
                    try {
                        result.add(type.getDeclaredField(component.getName()));
                    } catch (NoSuchFieldException e) {
                        throw new LitexlMapperException("Field not found: " + component.getName(), e);
                    }
                }
            }
        } else {
            for (Field field : type.getDeclaredFields()) {
                if (List.class.isAssignableFrom(field.getType())) {
                    result.add(field);
                }
            }
        }
        return result;
    }

    private Object readMultiRegionSheet(Sheet sheet, Class<?> type, List<Field> listFields) {
        Object[] values = new Object[listFields.size()];

        // Collect all header patterns for each list field
        List<Set<String>> allHeaders = new ArrayList<>();
        for (Field field : listFields) {
            Class<?> rowType = ReflectionHelper.getGenericListType(field);
            allHeaders.add(extractHeadersFromType(rowType));
        }

        for (int i = 0; i < listFields.size(); i++) {
            Field field = listFields.get(i);
            Class<?> rowType = ReflectionHelper.getGenericListType(field);
            Set<String> headers = allHeaders.get(i);

            // Combine other headers for end detection
            Set<String> otherHeaders = new HashSet<>();
            for (int j = 0; j < allHeaders.size(); j++) {
                if (j != i) {
                    otherHeaders.addAll(allHeaders.get(j));
                }
            }

            RegionDetector.Region region = RegionDetector.detectRegionWithEnd(
                    sheet, headers, otherHeaders.isEmpty() ? emptySet() : otherHeaders, 0);
            if (region != null) {
                values[i] = readRowsInRegion(sheet, rowType, region);
            } else {
                values[i] = List.of();
            }
        }

        return ReflectionHelper.createInstance(type, values);
    }

    private Set<String> extractHeadersFromType(Class<?> rowType) {
        Set<String> headers = new HashSet<>();
        if (rowType.isRecord()) {
            for (RecordComponent component : rowType.getRecordComponents()) {
                LitexlColumn ann = component.getAnnotation(LitexlColumn.class);
                if (ann != null && !ann.header().isEmpty()) {
                    headers.add(ann.header());
                }
            }
        } else {
            List<Field> columnFields = ReflectionHelper.getAnnotatedFields(rowType, LitexlColumn.class);
            for (Field field : columnFields) {
                LitexlColumn ann = field.getAnnotation(LitexlColumn.class);
                if (!ann.header().isEmpty()) {
                    headers.add(ann.header());
                }
            }
        }
        return headers;
    }

    private List<Object> readRowsInRegion(Sheet sheet, Class<?> rowType, RegionDetector.Region region) {
        validateRowAnnotation(rowType);
        Map<String, Integer> columnMap = buildColumnMapFromHeaderRow(sheet, region.headerRow());
        List<Object> result = new ArrayList<>();

        sheet.forEachRow(row -> {
            int rowIndex = row.rowNum();
            if (rowIndex < region.dataStartRow() || rowIndex > region.dataEndRow()) {
                return true;
            }

            if (ReflectionHelper.isEmptyRow(row)) {
                return true;
            }

            Object obj = readRowWithColumnMap(row, rowType, columnMap);
            if (obj != null) {
                result.add(obj);
            }
            return true;
        });

        return result;
    }

    private Map<String, Integer> buildColumnMapFromHeaderRow(Sheet sheet, int headerRow) {
        Row row = sheet.getRow(headerRow);
        Map<String, Integer> headerMap = new HashMap<>();

        if (row != null) {
            for (Map.Entry<Integer, Cell> entry : row.cells().entrySet()) {
                Cell cell = entry.getValue();
                if (cell.type() == CellType.STRING) {
                    headerMap.put(cell.string(), entry.getKey());
                }
            }
        }

        return headerMap;
    }

    private @Nullable Object readRowWithColumnMap(Row row, Class<?> type, Map<String, Integer> columnMap) {
        if (ReflectionHelper.isRecord(type)) {
            return readRowAsRecordWithColumnMap(row, type, columnMap);
        } else {
            return readRowAsClassWithColumnMap(row, type, columnMap);
        }
    }

    private @Nullable Object readRowAsRecordWithColumnMap(Row row, Class<?> type, Map<String, Integer> columnMap) {
        RecordComponent[] components = type.getRecordComponents();
        Object[] args = new Object[components.length];

        for (int i = 0; i < components.length; i++) {
            RecordComponent component = components[i];
            LitexlColumn columnAnnotation = component.getAnnotation(LitexlColumn.class);

            if (columnAnnotation != null) {
                int colIndex = getColumnIndex(columnAnnotation, columnMap);
                Cell cell = row.getCell(colIndex);
                CellValue cellValue = cell != null ? cell.value() : new CellValue.Empty();
                args[i] = convertCellValue(cellValue, component.getType(), columnAnnotation.converter());
            } else {
                args[i] = getDefaultValue(component.getType());
            }
        }

        return ReflectionHelper.createInstance(type, args);
    }

    private @Nullable Object readRowAsClassWithColumnMap(Row row, Class<?> type, Map<String, Integer> columnMap) {
        Object instance = ReflectionHelper.createInstance(type, null);

        List<Field> columnFields = ReflectionHelper.getAnnotatedFields(type, LitexlColumn.class);
        for (Field field : columnFields) {
            LitexlColumn columnAnnotation = field.getAnnotation(LitexlColumn.class);
            int colIndex = getColumnIndex(columnAnnotation, columnMap);

            Cell cell = row.getCell(colIndex);
            CellValue cellValue = cell != null ? cell.value() : new CellValue.Empty();
            Object value = convertCellValue(cellValue, field.getType(), columnAnnotation.converter());
            ReflectionHelper.setFieldValue(instance, field, value);
        }

        return instance;
    }

    private int getColumnIndex(LitexlColumn annotation, Map<String, Integer> columnMap) {
        if (!annotation.header().isEmpty() && columnMap.containsKey(annotation.header())) {
            return columnMap.get(annotation.header());
        } else if (annotation.index() >= 0) {
            return annotation.index();
        }
        throw new LitexlMapperException("Cannot determine column index for header: " + annotation.header());
    }

    @SuppressWarnings("unchecked")
    private @Nullable Object convertCellValue(CellValue cellValue, Class<?> targetType,
                                              Class<? extends LitexlConverter<?>> converterClass) {
        // Check for custom converter
        if (converterClass != LitexlConverter.None.class) {
            try {
                LitexlConverter<Object> converter = (LitexlConverter<Object>) converterClass.getDeclaredConstructor().newInstance();
                return converter.fromCell(cellValue);
            } catch (ReflectiveOperationException e) {
                throw new LitexlMapperException("Failed to create converter: " + converterClass.getName(), e);
            }
        }

        // Use TypeConverter for basic types
        if (TypeConverter.isSupported(targetType)) {
            return TypeConverter.fromCell(cellValue, targetType, config.dateFormat());
        }

        throw new LitexlMapperException("Unsupported type: " + targetType.getName() + ". Use a custom converter.");
    }

    private @Nullable Object getDefaultValue(Class<?> type) {
        if (type.isPrimitive()) {
            if (type == int.class) {
                return 0;
            } else if (type == long.class) {
                return 0L;
            } else if (type == double.class) {
                return 0.0;
            } else if (type == float.class) {
                return 0.0f;
            } else if (type == boolean.class) {
                return false;
            } else if (type == byte.class) {
                return (byte) 0;
            } else if (type == short.class) {
                return (short) 0;
            } else if (type == char.class) {
                return '\0';
            }
        }
        return null;
    }
}
