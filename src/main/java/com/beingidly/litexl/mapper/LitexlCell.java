package com.beingidly.litexl.mapper;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Maps a field to a specific cell by row and column index.
 */
@Target({ElementType.FIELD, ElementType.RECORD_COMPONENT})
@Retention(RetentionPolicy.RUNTIME)
public @interface LitexlCell {
    /**
     * The row index (zero-based).
     * @return the row index
     */
    int row();
    /**
     * The column index (zero-based).
     * @return the column index
     */
    int column();
    /**
     * Custom converter class.
     * @return the converter class
     */
    Class<? extends LitexlConverter<?>> converter() default LitexlConverter.None.class;
}
