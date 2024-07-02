package org.celper;


import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface CellFormat {
    Type builtinFormat() default Type.GENERAL;

    String customFormat() default "";

    enum Type {
        GENERAL("General"),
        GENERAL_NUMBER("0"),
        DECIMAL("0.00"),
        THOUSAND_SEPARATOR("#,##0"),
        ACCOUNTING("_-₩* #,##0_-;-₩* #,##0_-;_-₩* \"-\"_-;_-@_-"),
        SIMPLE_DATE("yyyy-mm-dd"),
        PERCENT("0%"),
        NUMBER_TO_KOREAN("[DBNum4]");
        public final String cellFormat;

        Type(String cellFormat) {
            this.cellFormat = cellFormat;
        }

        public String getCellFormat() {
            return cellFormat;
        }
    }


}
