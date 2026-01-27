package com.beingidly.litexl.mapper;

import com.beingidly.litexl.ErrorCode;
import com.beingidly.litexl.LitexlException;

public class LitexlMapperException extends LitexlException {

    public LitexlMapperException(String message) {
        super(ErrorCode.MAPPER_ERROR, message);
    }

    public LitexlMapperException(String message, Throwable cause) {
        super(ErrorCode.MAPPER_ERROR, message, cause);
    }
}
