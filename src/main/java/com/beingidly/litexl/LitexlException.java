package com.beingidly.litexl;

/**
 * Base exception for all litexl errors.
 */
public class LitexlException extends RuntimeException {

    private final ErrorCode code;

    public LitexlException(ErrorCode code, String message) {
        super(message);
        this.code = code;
    }

    public LitexlException(ErrorCode code, String message, Throwable cause) {
        super(message, cause);
        this.code = code;
    }

    public ErrorCode code() {
        return code;
    }
}
