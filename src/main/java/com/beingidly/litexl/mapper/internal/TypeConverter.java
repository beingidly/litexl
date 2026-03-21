package com.beingidly.litexl.mapper.internal;

import com.beingidly.litexl.CellValue;
import org.jspecify.annotations.Nullable;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Set;

/**
 * Converts between cell values and standard Java types.
 */
public final class TypeConverter {

    private static final Set<Class<?>> SUPPORTED_TYPES = Set.of(
        String.class,
        int.class, Integer.class,
        long.class, Long.class,
        double.class, Double.class,
        float.class, Float.class,
        boolean.class, Boolean.class,
        LocalDateTime.class,
        LocalDate.class
    );

    private TypeConverter() {}

    /**
     * Returns true if the given type is supported for conversion.
     *
     * @param type the type to check
     * @return true if supported
     */
    public static boolean isSupported(Class<?> type) {
        return SUPPORTED_TYPES.contains(type);
    }

    /**
     * Converts a cell value to the target type using default date format.
     *
     * @param <T> the target type
     * @param value the cell value
     * @param type the target class
     * @return the converted value, or null
     */
    public static <T> @Nullable T fromCell(CellValue value, Class<T> type) {
        return fromCell(value, type, null);
    }

    /**
     * Converts a cell value to the target type with an optional date format.
     *
     * @param <T> the target type
     * @param value the cell value
     * @param type the target class
     * @param dateFormat the date format pattern, or null
     * @return the converted value, or null
     */
    @SuppressWarnings("unchecked")
    public static <T> @Nullable T fromCell(CellValue value, Class<T> type, @Nullable String dateFormat) {
        if (value instanceof CellValue.Empty) {
            return (T) getDefaultValueForType(type);
        }

        if (type == String.class) {
            return (T) fromCellToString(value, dateFormat);
        }
        if (type == int.class || type == Integer.class) {
            return (T) fromCellToInt(value);
        }
        if (type == long.class || type == Long.class) {
            return (T) fromCellToLong(value);
        }
        if (type == double.class || type == Double.class) {
            return (T) fromCellToDouble(value);
        }
        if (type == float.class || type == Float.class) {
            return (T) fromCellToFloat(value);
        }
        if (type == boolean.class || type == Boolean.class) {
            return (T) fromCellToBoolean(value);
        }
        if (type == LocalDateTime.class) {
            return (T) fromCellToLocalDateTime(value);
        }
        if (type == LocalDate.class) {
            return (T) fromCellToLocalDate(value);
        }

        throw new IllegalArgumentException("Unsupported type: " + type);
    }

    /**
     * Converts a Java value to a cell value.
     *
     * @param value the Java value, or null
     * @param type the source type
     * @return the cell value
     */
    public static CellValue toCell(@Nullable Object value, Class<?> type) {
        if (value == null) {
            return new CellValue.Empty();
        }

        if (type == String.class) {
            return new CellValue.Text((String) value);
        }
        if (type == int.class || type == Integer.class) {
            return new CellValue.Number((Integer) value);
        }
        if (type == long.class || type == Long.class) {
            return new CellValue.Number((Long) value);
        }
        if (type == double.class || type == Double.class) {
            return new CellValue.Number((Double) value);
        }
        if (type == float.class || type == Float.class) {
            return new CellValue.Number((Float) value);
        }
        if (type == boolean.class || type == Boolean.class) {
            return new CellValue.Bool((Boolean) value);
        }
        if (type == LocalDateTime.class) {
            return new CellValue.Date((LocalDateTime) value);
        }
        if (type == LocalDate.class) {
            return new CellValue.Date(((LocalDate) value).atStartOfDay());
        }

        throw new IllegalArgumentException("Unsupported type: " + type);
    }

    private static @Nullable String fromCellToString(CellValue value, @Nullable String dateFormat) {
        return switch (value) {
            case CellValue.Text t -> t.value();
            case CellValue.Number n -> String.valueOf(n.value());
            case CellValue.Bool b -> String.valueOf(b.value());
            case CellValue.Date d -> {
                if (dateFormat != null) {
                    yield d.value().format(DateTimeFormatter.ofPattern(dateFormat));
                }
                yield d.value().toString();
            }
            default -> null;
        };
    }

    private static @Nullable Integer fromCellToInt(CellValue value) {
        return value instanceof CellValue.Number(double v) ? (int) v : null;
    }

    private static @Nullable Long fromCellToLong(CellValue value) {
        return value instanceof CellValue.Number(double v) ? (long) v : null;
    }

    private static @Nullable Double fromCellToDouble(CellValue value) {
        return value instanceof CellValue.Number(double v) ? v : null;
    }

    private static @Nullable Float fromCellToFloat(CellValue value) {
        return value instanceof CellValue.Number(double v) ? (float) v : null;
    }

    private static @Nullable Boolean fromCellToBoolean(CellValue value) {
        return value instanceof CellValue.Bool(boolean v) ? v : null;
    }

    private static @Nullable LocalDateTime fromCellToLocalDateTime(CellValue value) {
        return value instanceof CellValue.Date(LocalDateTime v) ? v : null;
    }

    private static @Nullable LocalDate fromCellToLocalDate(CellValue value) {
        return value instanceof CellValue.Date(LocalDateTime v) ? v.toLocalDate() : null;
    }

    private static @Nullable Object getDefaultValueForType(Class<?> type) {
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
        }
        return null;  // wrapper types and other types return null
    }
}
