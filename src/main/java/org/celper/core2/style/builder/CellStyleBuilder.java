package org.celper.core2.style.builder;

import org.apache.poi.ss.usermodel.*;

public class CellStyleBuilder {
    private final CellStyle cellStyle;
    private final Font font;

    public CellStyleBuilder(CellStyle cellStyle, Font font) {
        this.font = font;
        this.cellStyle = cellStyle;
    }

    public CellStyleBuilder setAllOfBorder(BorderStyle borderStyle) {
        this.cellStyle.setBorderBottom(borderStyle);
        this.cellStyle.setBorderTop(borderStyle);
        this.cellStyle.setBorderLeft(borderStyle);
        this.cellStyle.setBorderRight(borderStyle);
        return this;
    }

    public CellStyleBuilder setAlignment(HorizontalAlignment alignment) {
        this.cellStyle.setAlignment(alignment);
        return this;
    }

    public CellStyleBuilder setVerticalAlignment(VerticalAlignment verticalAlignment) {
        this.cellStyle.setVerticalAlignment(verticalAlignment);
        return this;
    }

    public CellStyleBuilder isWrapText(boolean b) {
        this.cellStyle.setWrapText(b);
        return this;
    }

    public CellStyleBuilder setFillBackgroundColor(short colorIndex, FillPatternType patternType) {
        this.cellStyle.setFillForegroundColor(colorIndex);
        this.cellStyle.setFillPattern(patternType);
        return this;
    }
    public CellStyleBuilder setFontHeightInPoints(short height) {
        this.font.setFontHeightInPoints(height);
        this.cellStyle.setFont(this.font);
        return this;
    }

    public CellStyleBuilder setFontName(String fontName) {
        this.font.setFontName(fontName);
        this.cellStyle.setFont(this.font);
        return this;
    }

    public CellStyleBuilder isBold(boolean b) {
        this.font.setBold(b);
        this.cellStyle.setFont(this.font);
        return this;
    }
}
