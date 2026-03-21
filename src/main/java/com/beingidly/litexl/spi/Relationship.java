package com.beingidly.litexl.spi;

/**
 * Represents an OPC relationship entry.
 *
 * @param id     the relationship id (e.g. "rId1")
 * @param type   the relationship type URI
 * @param target the target part path
 */
public record Relationship(String id, String type, String target) {
}
