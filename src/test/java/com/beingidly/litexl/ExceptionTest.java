package com.beingidly.litexl;

import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for exception classes to improve coverage.
 */
class ExceptionTest {

    // === LitexlException tests ===

    @Test
    void litexlException_withMessage() {
        var ex = new LitexlException(ErrorCode.IO_ERROR, "test message");
        assertEquals(ErrorCode.IO_ERROR, ex.code());
        assertTrue(ex.getMessage().contains("test message"));
    }

    @Test
    void litexlException_withCause() {
        var cause = new RuntimeException("cause");
        var ex = new LitexlException(ErrorCode.IO_ERROR, "test", cause);
        assertEquals(cause, ex.getCause());
        assertEquals(ErrorCode.IO_ERROR, ex.code());
    }

    @Test
    void litexlException_allErrorCodes() {
        // Test each error code to ensure they're all usable
        for (ErrorCode code : ErrorCode.values()) {
            var ex = new LitexlException(code, "test with " + code);
            assertEquals(code, ex.code());
        }
    }

    // === InvalidPasswordException tests ===

    @Test
    void invalidPasswordException_defaultConstructor() {
        var ex = new InvalidPasswordException();
        assertEquals(ErrorCode.INVALID_PASSWORD, ex.code());
        assertNotNull(ex.getMessage());
    }

    @Test
    void invalidPasswordException_withMessage() {
        var ex = new InvalidPasswordException("wrong password");
        assertEquals(ErrorCode.INVALID_PASSWORD, ex.code());
        assertTrue(ex.getMessage().contains("wrong password"));
    }

    @Test
    void invalidPasswordException_withCause() {
        var cause = new RuntimeException("inner cause");
        var ex = new InvalidPasswordException("invalid password", cause);
        assertEquals(ErrorCode.INVALID_PASSWORD, ex.code());
        assertEquals(cause, ex.getCause());
    }

    // === CorruptFileException tests ===

    @Test
    void corruptFileException_withDetail() {
        var ex = new CorruptFileException("corrupt data");
        assertEquals(ErrorCode.FILE_CORRUPT, ex.code());
        assertTrue(ex.getMessage().contains("corrupt"));
        assertTrue(ex.getMessage().contains("Corrupt file"));
    }

    @Test
    void corruptFileException_withCause() {
        var cause = new RuntimeException("inner");
        var ex = new CorruptFileException("corrupt file", cause);
        assertEquals(ErrorCode.FILE_CORRUPT, ex.code());
        assertEquals(cause, ex.getCause());
        assertTrue(ex.getMessage().contains("corrupt"));
    }

    @Test
    void corruptFileException_differentDetails() {
        var ex1 = new CorruptFileException("missing header");
        var ex2 = new CorruptFileException("invalid checksum");

        assertTrue(ex1.getMessage().contains("missing header"));
        assertTrue(ex2.getMessage().contains("invalid checksum"));
    }
}
