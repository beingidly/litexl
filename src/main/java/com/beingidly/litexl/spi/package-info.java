/**
 * Service Provider Interface (SPI) for extending LiteXL's XLSX read/write process.
 *
 * <p>Extension modules (e.g. {@code litexl-chart}) implement {@link WriteExtension}
 * and/or {@link ReadExtension}, registered via {@link java.util.ServiceLoader}.
 */
@NullMarked
package com.beingidly.litexl.spi;

import org.jspecify.annotations.NullMarked;
