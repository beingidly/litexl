package com.beingidly.litexl.mapper.internal;

import com.beingidly.litexl.Cell;
import com.beingidly.litexl.CellType;
import com.beingidly.litexl.Row;
import com.beingidly.litexl.mapper.LitexlMapperException;
import org.jspecify.annotations.Nullable;

import java.lang.annotation.Annotation;
import java.lang.reflect.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Reflection utilities for the object mapper.
 */
public final class ReflectionHelper {

    private ReflectionHelper() {}

    /**
     * Returns true if the given class is a record type.
     *
     * @param clazz the class to check
     * @return true if the class is a record
     */
    public static boolean isRecord(Class<?> clazz) {
        return clazz.isRecord();
    }

    /**
     * Creates an instance of the given class.
     *
     * @param <T> the type to create
     * @param clazz the class
     * @param args constructor arguments (used for records)
     * @return a new instance
     */
    public static <T> T createInstance(Class<T> clazz, Object[] args) {
        try {
            if (clazz.isRecord()) {
                return createRecordInstance(clazz, args);
            } else {
                return createClassInstance(clazz);
            }
        } catch (ReflectiveOperationException e) {
            throw new LitexlMapperException("Failed to create instance of " + clazz.getName(), e);
        }
    }

    private static <T> T createRecordInstance(Class<T> clazz, Object[] args) throws ReflectiveOperationException {
        var components = clazz.getRecordComponents();
        var types = new Class<?>[components.length];
        for (int i = 0; i < components.length; i++) {
            types[i] = components[i].getType();
        }
        var constructor = clazz.getDeclaredConstructor(types);
        constructor.setAccessible(true);
        return constructor.newInstance(args);
    }

    private static <T> T createClassInstance(Class<T> clazz) throws ReflectiveOperationException {
        var constructor = clazz.getDeclaredConstructor();
        constructor.setAccessible(true);
        return constructor.newInstance();
    }

    /**
     * Returns fields annotated with the given annotation.
     *
     * @param clazz the class to inspect
     * @param annotation the annotation to look for
     * @return list of annotated fields
     */
    public static List<Field> getAnnotatedFields(Class<?> clazz, Class<? extends Annotation> annotation) {
        var result = new ArrayList<Field>();

        if (clazz.isRecord()) {
            for (var component : clazz.getRecordComponents()) {
                if (component.isAnnotationPresent(annotation)) {
                    try {
                        var field = clazz.getDeclaredField(component.getName());
                        result.add(field);
                    } catch (NoSuchFieldException e) {
                        throw new LitexlMapperException("Field not found: " + component.getName(), e);
                    }
                }
            }
        } else {
            for (var field : clazz.getDeclaredFields()) {
                if (field.isAnnotationPresent(annotation)) {
                    result.add(field);
                }
            }
        }

        return result;
    }

    /**
     * Sets a field value on the given object.
     *
     * @param obj the target object
     * @param field the field to set
     * @param value the value to assign
     */
    public static void setFieldValue(Object obj, Field field, @Nullable Object value) {
        try {
            field.setAccessible(true);
            field.set(obj, value);
        } catch (IllegalAccessException e) {
            throw new LitexlMapperException("Failed to set field value: " + field.getName(), e);
        }
    }

    /**
     * Gets a field value from the given object.
     *
     * @param obj the source object
     * @param field the field to read
     * @return the field value, or null
     */
    public static @Nullable Object getFieldValue(Object obj, Field field) {
        try {
            field.setAccessible(true);
            return field.get(obj);
        } catch (IllegalAccessException e) {
            throw new LitexlMapperException("Failed to get field value: " + field.getName(), e);
        }
    }

    /**
     * Returns the generic element type of a List field.
     *
     * @param field the list field
     * @return the element type
     */
    public static Class<?> getGenericListType(Field field) {
        var genericType = field.getGenericType();
        if (genericType instanceof ParameterizedType pt) {
            var typeArgs = pt.getActualTypeArguments();
            if (typeArgs.length > 0 && typeArgs[0] instanceof Class<?> c) {
                return c;
            }
        }
        throw new LitexlMapperException("Cannot determine generic type for field: " + field.getName());
    }

    /**
     * Returns true if the field type is assignable to List.
     *
     * @param field the field to check
     * @return true if the field is a list
     */
    public static boolean isList(Field field) {
        return List.class.isAssignableFrom(field.getType());
    }

    /**
     * Returns the first type argument from a parameterized type.
     *
     * @param genericType the generic type
     * @return the first type argument
     */
    public static Class<?> getGenericType(Type genericType) {
        if (genericType instanceof ParameterizedType pt) {
            Type[] typeArgs = pt.getActualTypeArguments();
            if (typeArgs.length > 0 && typeArgs[0] instanceof Class<?> c) {
                return c;
            }
        }
        throw new LitexlMapperException("Cannot determine generic type for: " + genericType);
    }

    /**
     * Returns true if the row contains only empty cells.
     *
     * @param row the row to check
     * @return true if all cells are empty
     */
    public static boolean isEmptyRow(Row row) {
        for (Cell cell : row.cells().values()) {
            if (cell.type() != CellType.EMPTY) {
                return false;
            }
        }
        return true;
    }
}
