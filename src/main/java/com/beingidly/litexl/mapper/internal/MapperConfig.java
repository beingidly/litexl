package com.beingidly.litexl.mapper.internal;

import com.beingidly.litexl.mapper.NullStrategy;

public record MapperConfig(
    String dateFormat,
    NullStrategy nullStrategy
) {
    public static MapperConfig defaults() {
        return new MapperConfig("yyyy-MM-dd HH:mm:ss", NullStrategy.SKIP);
    }
}
