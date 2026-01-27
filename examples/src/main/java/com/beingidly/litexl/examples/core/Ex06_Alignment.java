package com.beingidly.litexl.examples.core;

import com.beingidly.litexl.Sheet;
import com.beingidly.litexl.Workbook;
import com.beingidly.litexl.examples.util.ExampleUtils;
import com.beingidly.litexl.style.BorderStyle;
import com.beingidly.litexl.style.HAlign;
import com.beingidly.litexl.style.Style;
import com.beingidly.litexl.style.VAlign;

import java.nio.file.Path;

/**
 * Example: Text alignment and wrapping.
 *
 * This example demonstrates:
 * - Horizontal alignment (left, center, right)
 * - Vertical alignment (top, middle, bottom)
 * - Text wrapping
 */
public class Ex06_Alignment {

    public static void main(String[] args) {
        Path outputPath = ExampleUtils.tempFile("ex06_alignment.xlsx");

        try (Workbook workbook = Workbook.create()) {
            Sheet sheet = workbook.addSheet("Alignment");

            // Horizontal alignments
            int leftStyle = workbook.addStyle(Style.builder()
                .align(HAlign.LEFT, VAlign.MIDDLE)
                .border(BorderStyle.THIN, 0xFF000000)
                .build());

            int centerStyle = workbook.addStyle(Style.builder()
                .align(HAlign.CENTER, VAlign.MIDDLE)
                .border(BorderStyle.THIN, 0xFF000000)
                .build());

            int rightStyle = workbook.addStyle(Style.builder()
                .align(HAlign.RIGHT, VAlign.MIDDLE)
                .border(BorderStyle.THIN, 0xFF000000)
                .build());

            // Vertical alignments
            int topStyle = workbook.addStyle(Style.builder()
                .align(HAlign.CENTER, VAlign.TOP)
                .border(BorderStyle.THIN, 0xFF000000)
                .build());

            int middleStyle = workbook.addStyle(Style.builder()
                .align(HAlign.CENTER, VAlign.MIDDLE)
                .border(BorderStyle.THIN, 0xFF000000)
                .build());

            int bottomStyle = workbook.addStyle(Style.builder()
                .align(HAlign.CENTER, VAlign.BOTTOM)
                .border(BorderStyle.THIN, 0xFF000000)
                .build());

            // Text wrapping
            int wrapStyle = workbook.addStyle(Style.builder()
                .wrap(true)
                .align(HAlign.LEFT, VAlign.TOP)
                .border(BorderStyle.THIN, 0xFF000000)
                .build());

            // Header style
            int headerStyle = workbook.addStyle(Style.builder()
                .bold(true)
                .align(HAlign.CENTER, VAlign.MIDDLE)
                .fill(0xFFD9E1F2)
                .border(BorderStyle.THIN, 0xFF000000)
                .build());

            // Horizontal alignment demo
            sheet.cell(0, 0).set("Horizontal Alignment").style(headerStyle);
            sheet.cell(1, 0).set("Left aligned").style(leftStyle);
            sheet.cell(2, 0).set("Center aligned").style(centerStyle);
            sheet.cell(3, 0).set("Right aligned").style(rightStyle);

            // Vertical alignment demo (use taller rows)
            sheet.cell(0, 1).set("Vertical Alignment").style(headerStyle);
            sheet.cell(1, 1).set("Top").style(topStyle);
            sheet.cell(2, 1).set("Middle").style(middleStyle);
            sheet.cell(3, 1).set("Bottom").style(bottomStyle);

            // Set row heights for vertical alignment visibility
            sheet.row(1).height(40);
            sheet.row(2).height(40);
            sheet.row(3).height(40);

            // Text wrapping demo
            sheet.cell(0, 2).set("Text Wrapping").style(headerStyle);
            sheet.cell(1, 2).set("This is a long text that will wrap to multiple lines when the column width is not enough.").style(wrapStyle);

            // Set column widths
            sheet.setColumnWidth(0, 20);
            sheet.setColumnWidth(1, 20);
            sheet.setColumnWidth(2, 25);

            workbook.save(outputPath);
        }

        ExampleUtils.printCreated(outputPath);
    }
}
