package com.beingidly.litexl.mapper;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Maps a field to a column by index or header name.
 */
@Target({ElementType.FIELD, ElementType.RECORD_COMPONENT})
@Retention(RetentionPolicy.RUNTIME)
public @interface LitexlColumn {
    /**
     * The column index (zero-based), or -1 to use header matching.
     * @return the column index
     */
    int index() default -1;
    /**
     * The header name to match, or empty to use index.
     * @return the header name
     */
    String header() default "";
    /**
     * Custom converter class.
     * @return the converter class
     */
    Class<? extends LitexlConverter<?>> converter() default LitexlConverter.None.class;
}
