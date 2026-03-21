package com.beingidly.litexl.mapper;

import com.beingidly.litexl.style.Style;

/**
 * Provides a cell style for mapped fields.
 */
public interface LitexlStyleProvider {
    /**
     * Returns the style to apply.
     *
     * @return the cell style
     */
    Style provide();
}
