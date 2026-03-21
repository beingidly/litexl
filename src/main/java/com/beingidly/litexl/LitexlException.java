package com.beingidly.litexl;

/**
 * Base exception for all litexl errors.
 */
public class LitexlException extends RuntimeException {

    /** The error code. */
    private final ErrorCode code;

    /**
     * Creates a litexl exception with the given error code and message.
     *
     * @param code the error code
     * @param message the detail message
     */
    public LitexlException(ErrorCode code, String message) {
        super(message);
        this.code = code;
    }

    /**
     * Creates a litexl exception with the given error code, message, and cause.
     *
     * @param code the error code
     * @param message the detail message
     * @param cause the underlying cause
     */
    public LitexlException(ErrorCode code, String message, Throwable cause) {
        super(message, cause);
        this.code = code;
    }

    /**
     * Returns the error code.
     *
     * @return the error code
     */
    public ErrorCode code() {
        return code;
    }
}
