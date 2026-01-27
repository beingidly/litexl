package com.beingidly.litexl;

/**
 * Utility class for converting between cell references (A1) and coordinates (0, 0).
 */
final class CellRefUtil {

    private CellRefUtil() {}

    /**
     * Converts a column index (0-based) to Excel column letters.
     * 0 -> A, 1 -> B, ..., 25 -> Z, 26 -> AA, etc.
     */
    public static String colToLetters(int col) {
        if (col < 0) {
            throw new IllegalArgumentException("Column index must be non-negative: " + col);
        }

        StringBuilder sb = new StringBuilder();
        int c = col + 1; // Convert to 1-based

        while (c > 0) {
            c--;
            sb.insert(0, (char) ('A' + (c % 26)));
            c /= 26;
        }

        return sb.toString();
    }

    /**
     * Converts Excel column letters to a column index (0-based).
     * A -> 0, B -> 1, ..., Z -> 25, AA -> 26, etc.
     */
    public static int lettersToCol(String letters) {
        if (letters.isEmpty()) {
            throw new IllegalArgumentException("Column letters cannot be empty");
        }

        int col = 0;
        for (int i = 0; i < letters.length(); i++) {
            char ch = letters.charAt(i);
            if (ch < 'A' || ch > 'Z') {
                throw new IllegalArgumentException("Invalid column letter: " + ch);
            }
            col = col * 26 + (ch - 'A' + 1);
        }

        return col - 1; // Convert to 0-based
    }

    /**
     * Converts row and column indices to a cell reference like "A1".
     */
    public static String toRef(int row, int col) {
        return colToLetters(col) + (row + 1);
    }

    /**
     * Converts row and column indices to an absolute cell reference like "$A$1".
     */
    public static String toAbsoluteRef(int row, int col) {
        return "$" + colToLetters(col) + "$" + (row + 1);
    }

    /**
     * Parses a cell reference like "A1" to [row, col] (0-based).
     */
    public static int[] parseRef(String ref) {
        if (ref.isEmpty()) {
            throw new IllegalArgumentException("Cell reference cannot be empty");
        }

        // Find where letters end and digits begin
        int i = 0;
        while (i < ref.length() && Character.isLetter(ref.charAt(i))) {
            i++;
        }

        if (i == 0) {
            throw new IllegalArgumentException("Invalid cell reference (no column letters): " + ref);
        }
        if (i == ref.length()) {
            throw new IllegalArgumentException("Invalid cell reference (no row number): " + ref);
        }

        String colPart = ref.substring(0, i).toUpperCase();
        String rowPart = ref.substring(i);

        int col = lettersToCol(colPart);

        int row;
        try {
            row = Integer.parseInt(rowPart) - 1; // Convert to 0-based
        } catch (NumberFormatException e) {
            throw new IllegalArgumentException("Invalid row number in cell reference: " + ref);
        }

        if (row < 0) {
            throw new IllegalArgumentException("Row number must be positive: " + ref);
        }

        return new int[]{row, col};
    }
}
