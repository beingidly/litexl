package com.beingidly.litexl.mapper;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Associates a style provider with a field.
 */
@Target({ElementType.FIELD, ElementType.RECORD_COMPONENT})
@Retention(RetentionPolicy.RUNTIME)
public @interface LitexlStyle {
    /**
     * The style provider class.
     * @return the style provider class
     */
    Class<? extends LitexlStyleProvider> value();
}
