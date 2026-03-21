package com.beingidly.litexl.mapper.internal;

import com.beingidly.litexl.mapper.NullStrategy;

/**
 * Configuration for the object mapper.
 *
 * @param dateFormat the date format pattern
 * @param nullStrategy the null handling strategy
 */
public record MapperConfig(
    String dateFormat,
    NullStrategy nullStrategy
) {
    /**
     * Returns the default configuration.
     *
     * @return the default mapper config
     */
    public static MapperConfig defaults() {
        return new MapperConfig("yyyy-MM-dd HH:mm:ss", NullStrategy.SKIP);
    }
}
