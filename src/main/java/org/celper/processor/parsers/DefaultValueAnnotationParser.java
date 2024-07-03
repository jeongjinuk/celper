package org.celper.processor.parsers;

import org.celper.DefaultValue;
import org.celper.processor.FieldAnnotationMetaData;
import org.celper.processor.Parser;

import java.util.Objects;

public class DefaultValueAnnotationParser implements Parser<FieldAnnotationMetaData, FieldAnnotationMetaData> {
    @Override
    public FieldAnnotationMetaData parse(FieldAnnotationMetaData fieldAnnotationMetaData) {
        DefaultValue annotation = fieldAnnotationMetaData.getField().getAnnotation(DefaultValue.class);
        if (Objects.nonNull(annotation)){
            fieldAnnotationMetaData.setDefaultValue(annotation.value());
        }
        return fieldAnnotationMetaData;
    }
}
