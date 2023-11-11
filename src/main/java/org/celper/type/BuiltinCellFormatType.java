package org.celper.type;

public enum BuiltinCellFormatType {
    GENERAL("General"),
    GENERAL_NUMBER("0"),
    DECIMAL("0.00"),
    THOUSAND_SEPARATOR("#,##0"),
    ACCOUNTING("_-₩* #,##0_-;-₩* #,##0_-;_-₩* \"-\"_-;_-@_-"),
    SIMPLE_DATE("yyyy-mm-dd"),
    PERCENT("0%"),
    NUMBER_TO_KOREAN("[DBNum4]");
    public final String cellFormat;
    BuiltinCellFormatType(String cellFormat) {
        this.cellFormat = cellFormat;
    }

    public String getCellFormat() {
        return cellFormat;
    }
}
