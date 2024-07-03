package org.celper.processor.parsers;

import org.celper.CellFormat;
import org.celper.processor.FieldAnnotationMetaData;
import org.celper.processor.Parser;

public class CellFormatAnnotationParser implements Parser<FieldAnnotationMetaData, FieldAnnotationMetaData> {
    @Override
    public FieldAnnotationMetaData parse(FieldAnnotationMetaData fieldAnnotationMetaData) {
        CellFormat annotation = fieldAnnotationMetaData.getField().getAnnotation(CellFormat.class);
        String format = !"".equals(annotation.customFormat()) ? annotation.customFormat() : annotation.builtinFormat().getCellFormat();
        fieldAnnotationMetaData.setCellFormat(format);
        return fieldAnnotationMetaData;
    }
}
