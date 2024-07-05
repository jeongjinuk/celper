package org.celper.processor.parser;

import org.celper.CellFormat;
import org.celper.Column;
import org.celper.ColumnStyle;
import org.celper.DefaultValue;
import org.celper.core.style.CellStyleConfigurer;
import org.celper.processor.meta.FieldAnnotationMetaData;

import javax.lang.model.element.VariableElement;
import java.lang.annotation.Annotation;
import java.util.Objects;
import java.util.function.BiConsumer;

public enum FieldAnnotationHandler implements AnnotationHandler<VariableElement, FieldAnnotationMetaData> {
    COLUMN(Column.class, (annotation, metaData) -> {
        Column column = (Column) annotation;
        metaData.setHeaderName(column.value());
        metaData.setPriority(column.priority());
    }),

    DEFAULT_VALUE(DefaultValue.class, (annotation, metaData) -> {
        DefaultValue defaultValue = (DefaultValue) annotation;
        metaData.setDefaultValue(defaultValue.value());
    }),

    CELL_FORMAT(CellFormat.class, (annotation, metaData) -> {
        CellFormat cellFormat = (CellFormat) annotation;
        metaData.setCellFormat(!"".equals(cellFormat.customFormat()) ? cellFormat.customFormat() : cellFormat.builtinFormat().getCellFormat());
    }),

    COLUMN_STYLE(ColumnStyle.class, (annotation, metaData) -> {
        ColumnStyle columnStyle = (ColumnStyle) annotation;
        metaData.setHeaderStyleConfigurer((Class<CellStyleConfigurer>) columnStyle.headerAreaStyle());
        metaData.setDataStyleConfigurer((Class<CellStyleConfigurer>) columnStyle.dataAreaStyle());
    });

    private final Class<? extends Annotation> type;
    private final BiConsumer<Annotation, FieldAnnotationMetaData> ifPresent;

    FieldAnnotationHandler(Class<? extends Annotation> type, BiConsumer<Annotation, FieldAnnotationMetaData> ifPresent) {
        this.type = type;
        this.ifPresent = ifPresent;
    }
    @Override
    public void ifPresent(VariableElement field,
                         FieldAnnotationMetaData metaData) {
        Annotation annotation = field.getAnnotation(this.type);
        if (Objects.nonNull(annotation))
            this.ifPresent.accept(annotation, metaData);
    }

}
