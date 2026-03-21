package com.beingidly.litexl.mapper.internal;

import com.beingidly.litexl.*;
import com.beingidly.litexl.mapper.LitexlColumn;
import org.jspecify.annotations.Nullable;

import java.util.*;

/**
 * Detects data regions within a sheet by scanning for header rows.
 */
public final class RegionDetector {

    private RegionDetector() {}

    /**
     * Represents a detected region within a sheet.
     *
     * @param headerRow the header row index
     * @param dataStartRow the first data row index
     * @param dataEndRow the last data row index
     */
    public record Region(int headerRow, int dataStartRow, int dataEndRow) {
        /**
         * Creates a region with no defined end row.
         *
         * @param headerRow the header row index
         * @param dataStartRow the first data row index
         */
        public Region(int headerRow, int dataStartRow) {
            this(headerRow, dataStartRow, Integer.MAX_VALUE);
        }
    }

    /**
     * Detects the first region matching the given headers.
     *
     * @param sheet the sheet to scan
     * @param headers the expected header names
     * @param startRow the row to start scanning from
     * @return the detected region, or null if not found
     */
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

    /**
     * Detects a region with a defined end row, bounded by the next region's headers.
     *
     * @param sheet the sheet to scan
     * @param headers the expected header names
     * @param nextHeaders the headers of the next region (used to find the end)
     * @param startRow the row to start scanning from
     * @return the detected region, or null if not found
     */
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
        final int[] lastDataRow = {dataStartRow - 1};
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

    /**
     * Extracts header names from annotation metadata on the given row type.
     *
     * @param rowType the row class annotated with {@link LitexlColumn}
     * @return the set of header names
     */
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
