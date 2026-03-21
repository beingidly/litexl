package com.beingidly.litexl.mapper;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Maps a type or field to a worksheet.
 */
@Target({ElementType.TYPE, ElementType.FIELD, ElementType.RECORD_COMPONENT})
@Retention(RetentionPolicy.RUNTIME)
public @interface LitexlSheet {
    /**
     * The sheet name, or empty to use index.
     * @return the sheet name
     */
    String name() default "";
    /**
     * The sheet index (zero-based), or -1 to use name.
     * @return the sheet index
     */
    int index() default -1;
    /**
     * The header row index (zero-based).
     * @return the header row index
     */
    int headerRow() default 0;
    /**
     * The data start row index (zero-based).
     * @return the data start row index
     */
    int dataStartRow() default 1;
    /**
     * The data start column index (zero-based).
     * @return the data start column index
     */
    int dataStartColumn() default 0;
    /**
     * The region detection strategy.
     * @return the region detection mode
     */
    RegionDetection regionDetection() default RegionDetection.NONE;
}
