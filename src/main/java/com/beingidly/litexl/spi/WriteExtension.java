package com.beingidly.litexl.spi;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.XmlWriter;

import java.io.IOException;
import java.util.List;

/**
 * SPI for extending the XLSX write process.
 *
 * <p>Implementations are discovered via {@link java.util.ServiceLoader}.
 * When {@code litexl-chart} (or another extension) is on the classpath,
 * its {@code WriteExtension} is automatically invoked during save.
 */
public interface WriteExtension {

    /**
     * Called during [Content_Types].xml writing to register additional content types.
     *
     * @param registry   the content type registry
     * @param sheetCount the number of sheets in the workbook
     * @param sheets     the sheets in the workbook
     */
    void contributeContentTypes(ContentTypeRegistry registry,
                                List<Sheet> sheets) throws IOException;

    /**
     * Called inside each sheet's XML, after dataValidations and before
     * the closing {@code </worksheet>} element.
     *
     * <p>Typically used to write {@code <drawing r:id="rId1"/>} elements.
     *
     * @param xml      the current sheet's XML writer
     * @param sheet    the sheet being written
     * @param sheetNum the 1-based sheet number
     */
    void writeSheetElements(XmlWriter xml, Sheet sheet, int sheetNum) throws IOException;

    /**
     * Returns relationships to include in the sheet's .rels file.
     *
     * <p>If the returned list is non-empty, a relationship file will be created
     * at {@code xl/worksheets/_rels/sheet{sheetNum}.xml.rels}.
     *
     * @param sheet    the sheet
     * @param sheetNum the 1-based sheet number
     * @return list of relationships, or empty list if none
     */
    List<Relationship> sheetRelationships(Sheet sheet, int sheetNum);

    /**
     * Called after all sheets are written to produce additional ZIP entries
     * (e.g. chart XML, drawing XML, media files).
     *
     * @param ctx    the write context for creating ZIP entries
     * @param sheets all sheets in the workbook
     */
    void writeEntries(WriteContext ctx, List<Sheet> sheets) throws IOException;
}
