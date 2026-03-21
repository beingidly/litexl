package com.beingidly.litexl;

/**
 * Thrown when an incorrect password is provided for an encrypted file.
 */
public class InvalidPasswordException extends LitexlException {

    /** Creates an invalid password exception with a default message. */
    public InvalidPasswordException() {
        super(ErrorCode.INVALID_PASSWORD, "Invalid password");
    }

    /**
     * Creates an invalid password exception with a custom message.
     *
     * @param message the detail message
     */
    public InvalidPasswordException(String message) {
        super(ErrorCode.INVALID_PASSWORD, message);
    }

    /**
     * Creates an invalid password exception with a custom message and cause.
     *
     * @param message the detail message
     * @param cause the underlying cause
     */
    public InvalidPasswordException(String message, Throwable cause) {
        super(ErrorCode.INVALID_PASSWORD, message);
        initCause(cause);
    }
}
