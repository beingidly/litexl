package com.beingidly.litexl;

/**
 * Thrown when an incorrect password is provided for an encrypted file.
 */
public class InvalidPasswordException extends LitexlException {

    public InvalidPasswordException() {
        super(ErrorCode.INVALID_PASSWORD, "Invalid password");
    }

    public InvalidPasswordException(String message) {
        super(ErrorCode.INVALID_PASSWORD, message);
    }

    public InvalidPasswordException(String message, Throwable cause) {
        super(ErrorCode.INVALID_PASSWORD, message);
        initCause(cause);
    }
}
