package com.beingidly.litexl.mapper.internal;

import com.beingidly.litexl.mapper.NullStrategy;
import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class MapperConfigTest {

    @Test
    void defaultConfig() {
        var config = MapperConfig.defaults();
        assertEquals("yyyy-MM-dd HH:mm:ss", config.dateFormat());
        assertEquals(NullStrategy.SKIP, config.nullStrategy());
    }

    @Test
    void customConfig() {
        var config = new MapperConfig("yyyy/MM/dd", NullStrategy.EMPTY_CELL);
        assertEquals("yyyy/MM/dd", config.dateFormat());
        assertEquals(NullStrategy.EMPTY_CELL, config.nullStrategy());
    }
}
