package org.celper.processor.parsers;

import org.celper.Column;
import org.celper.processor.FieldAnnotationMetaData;
import org.celper.processor.Parser;

public class ColumnAnnotationParser implements Parser<FieldAnnotationMetaData, FieldAnnotationMetaData> {
    @Override
    public FieldAnnotationMetaData parse(FieldAnnotationMetaData fieldAnnotationMetaData) {
        Column annotation = fieldAnnotationMetaData.getField().getAnnotation(Column.class);
        fieldAnnotationMetaData.setPriority(annotation.priority());
        fieldAnnotationMetaData.setHeaderName(annotation.value());
        return fieldAnnotationMetaData;
    }
}
