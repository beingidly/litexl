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
        for (var entry : sheet.rows().entrySet()) {
            int rowIndex = entry.getKey();
            if (rowIndex < startRow) {
                continue;
            }

            var row = entry.getValue();
            if (matchesHeaders(row, headers)) {
                return new Region(rowIndex, rowIndex + 1);
            }
        }
        return null;
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
        int lastDataRow = dataStartRow - 1;

        for (var entry : sheet.rows().entrySet()) {
            int rowIndex = entry.getKey();
            if (rowIndex < dataStartRow) {
                continue;
            }

            var row = entry.getValue();

            // Check if this row matches the next region's headers
            if (matchesHeaders(row, nextHeaders)) {
                return lastDataRow;
            }

            // Check if this is an empty row (potential region separator)
            if (ReflectionHelper.isEmptyRow(row)) {
                continue;
            }

            lastDataRow = rowIndex;
        }

        return lastDataRow;
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
