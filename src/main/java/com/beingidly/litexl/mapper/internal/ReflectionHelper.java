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

public final class ReflectionHelper {

    private ReflectionHelper() {}

    public static boolean isRecord(Class<?> clazz) {
        return clazz.isRecord();
    }

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

    public static void setFieldValue(Object obj, Field field, @Nullable Object value) {
        try {
            field.setAccessible(true);
            field.set(obj, value);
        } catch (IllegalAccessException e) {
            throw new LitexlMapperException("Failed to set field value: " + field.getName(), e);
        }
    }

    public static @Nullable Object getFieldValue(Object obj, Field field) {
        try {
            field.setAccessible(true);
            return field.get(obj);
        } catch (IllegalAccessException e) {
            throw new LitexlMapperException("Failed to get field value: " + field.getName(), e);
        }
    }

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

    public static boolean isList(Field field) {
        return List.class.isAssignableFrom(field.getType());
    }

    public static Class<?> getGenericType(Type genericType) {
        if (genericType instanceof ParameterizedType pt) {
            Type[] typeArgs = pt.getActualTypeArguments();
            if (typeArgs.length > 0 && typeArgs[0] instanceof Class<?> c) {
                return c;
            }
        }
        throw new LitexlMapperException("Cannot determine generic type for: " + genericType);
    }

    public static boolean isEmptyRow(Row row) {
        for (Cell cell : row.cells().values()) {
            if (cell.type() != CellType.EMPTY) {
                return false;
            }
        }
        return true;
    }
}
