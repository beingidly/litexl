package com.beingidly.litexl.mapper;

import com.beingidly.litexl.ErrorCode;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class LitexlMapperExceptionTest {

    @Test
    void createWithMessage() {
        var ex = new LitexlMapperException("test error");
        assertEquals("test error", ex.getMessage());
        assertEquals(ErrorCode.MAPPER_ERROR, ex.code());
    }

    @Test
    void createWithCause() {
        var cause = new RuntimeException("cause");
        var ex = new LitexlMapperException("test error", cause);
        assertEquals("test error", ex.getMessage());
        assertSame(cause, ex.getCause());
        assertEquals(ErrorCode.MAPPER_ERROR, ex.code());
    }
}
