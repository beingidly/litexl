package com.beingidly.litexl.chart.internal;

import com.beingidly.litexl.chart.Chart;
import com.beingidly.litexl.chart.ChartSeries;
import com.beingidly.litexl.chart.style.ChartFill;

import java.util.ArrayList;
import java.util.List;

/**
 * Collects media files (images) from chart fills.
 */
final class MediaWriter {

    record MediaEntry(String path, byte[] data, String mimeType) {}

    static List<MediaEntry> collectMedia(Chart chart, int baseMediaNum) {
        List<MediaEntry> entries = new ArrayList<>();
        int mediaNum = baseMediaNum;

        for (ChartSeries series : chart.series()) {
            if (series.fill() instanceof ChartFill.Picture pic) {
                mediaNum++;
                String ext = extensionFromMimeType(pic.mimeType());
                entries.add(new MediaEntry(
                    "xl/media/image" + mediaNum + "." + ext,
                    pic.imageData(),
                    pic.mimeType()));
            }
        }

        return entries;
    }

    private static String extensionFromMimeType(String mimeType) {
        return switch (mimeType) {
            case "image/png" -> "png";
            case "image/jpeg", "image/jpg" -> "jpeg";
            case "image/gif" -> "gif";
            case "image/bmp" -> "bmp";
            case "image/tiff" -> "tiff";
            default -> "png";
        };
    }

    private MediaWriter() {}
}
