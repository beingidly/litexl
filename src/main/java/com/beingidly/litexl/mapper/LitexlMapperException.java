package com.beingidly.litexl.mapper;

import com.beingidly.litexl.ErrorCode;
import com.beingidly.litexl.LitexlException;

/**
 * Thrown when an error occurs during object mapping.
 */
public class LitexlMapperException extends LitexlException {

    /**
     * Creates a mapper exception with the given message.
     *
     * @param message the detail message
     */
    public LitexlMapperException(String message) {
        super(ErrorCode.MAPPER_ERROR, message);
    }

    /**
     * Creates a mapper exception with the given message and cause.
     *
     * @param message the detail message
     * @param cause the underlying cause
     */
    public LitexlMapperException(String message, Throwable cause) {
        super(ErrorCode.MAPPER_ERROR, message, cause);
    }
}
