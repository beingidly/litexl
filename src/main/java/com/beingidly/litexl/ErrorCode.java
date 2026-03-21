package com.beingidly.litexl;

/**
 * Error codes for litexl exceptions.
 */
public enum ErrorCode {
    /** File was not found. */
    FILE_NOT_FOUND,
    /** File is corrupted or malformed. */
    FILE_CORRUPT,
    /** Incorrect password for encrypted file. */
    INVALID_PASSWORD,
    /** Unsupported file format. */
    UNSUPPORTED_FORMAT,
    /** Invalid argument provided. */
    INVALID_ARGUMENT,
    /** General I/O error. */
    IO_ERROR,
    /** XML parsing error. */
    XML_PARSE_ERROR,
    /** ZIP processing error. */
    ZIP_ERROR,
    /** Cryptography error. */
    CRYPTO_ERROR,
    /** Object mapper error. */
    MAPPER_ERROR
}
