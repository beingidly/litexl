package com.beingidly.litexl;

/**
 * Thrown when a file is corrupted or malformed.
 */
public class CorruptFileException extends LitexlException {

    public CorruptFileException(String detail) {
        super(ErrorCode.FILE_CORRUPT, "Corrupt file: " + detail);
    }

    public CorruptFileException(String detail, Throwable cause) {
        super(ErrorCode.FILE_CORRUPT, "Corrupt file: " + detail, cause);
    }
}
