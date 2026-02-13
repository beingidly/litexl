package com.beingidly.litexl.mapper.internal;

import com.beingidly.litexl.*;
import com.beingidly.litexl.mapper.LitexlColumn;
import org.jspecify.annotations.Nullable;

import java.util.*;

public final class RegionDetector {

    private RegionDetector() {}

    public record Region(int headerRow, int dataStartRow, int dataEndRow) {
        public Region(int headerRow, int dataStartRow) {
            this(headerRow, dataStartRow, Integer.MAX_VALUE);
        }
    }

    public static @Nullable Region detectRegion(Sheet sheet, Set<String> headers, int startRow) {
        final Region[] result = new Region[1];
        sheet.forEachRow(row -> {
            int rowIndex = row.rowNum();
            if (rowIndex < startRow) {
                return true;
            }
            if (matchesHeaders(row, headers)) {
                result[0] = new Region(rowIndex, rowIndex + 1);
                return false;
            }
            return true;
        });
        return result[0];
    }

    public static @Nullable Region detectRegionWithEnd(
            Sheet sheet,
            Set<String> headers,
            Set<String> nextHeaders,
            int startRow) {

        var region = detectRegion(sheet, headers, startRow);
        if (region == null) {
            return null;
        }

        int endRow = findEndRow(sheet, region.dataStartRow(), nextHeaders);
        return new Region(region.headerRow(), region.dataStartRow(), endRow);
    }

    private static boolean matchesHeaders(Row row, Set<String> expectedHeaders) {
        var actualHeaders = new HashSet<String>();
        for (var cell : row.cells().values()) {
            if (cell.type() == CellType.STRING) {
                actualHeaders.add(cell.string());
            }
        }
        return actualHeaders.containsAll(expectedHeaders);
    }

    private static int findEndRow(Sheet sheet, int dataStartRow, Set<String> nextHeaders) {
        final int[] lastDataRow = { dataStartRow - 1 };
        sheet.forEachRow(row -> {
            int rowIndex = row.rowNum();
            if (rowIndex < dataStartRow) {
                return true;
            }

            // Check if this row matches the next region's headers
            if (matchesHeaders(row, nextHeaders)) {
                return false;
            }

            // Check if this is an empty row (potential region separator)
            if (ReflectionHelper.isEmptyRow(row)) {
                return true;
            }

            lastDataRow[0] = rowIndex;
            return true;
        });

        return lastDataRow[0];
    }

    public static Set<String> extractHeaders(Class<?> rowType) {
        var fields = ReflectionHelper.getAnnotatedFields(rowType, LitexlColumn.class);
        var headers = new HashSet<String>();

        for (var field : fields) {
            var ann = field.getAnnotation(LitexlColumn.class);
            if (!ann.header().isEmpty()) {
                headers.add(ann.header());
            }
        }

        return headers;
    }
}
