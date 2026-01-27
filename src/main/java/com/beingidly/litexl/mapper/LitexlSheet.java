package com.beingidly.litexl.mapper;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.TYPE, ElementType.FIELD, ElementType.RECORD_COMPONENT})
@Retention(RetentionPolicy.RUNTIME)
public @interface LitexlSheet {
    String name() default "";
    int index() default -1;
    int headerRow() default 0;
    int dataStartRow() default 1;
    int dataStartColumn() default 0;
    RegionDetection regionDetection() default RegionDetection.NONE;
}
