package org.celper.processor.parser;

import org.celper.CellFormat;
import org.celper.Column;
import org.celper.ColumnStyle;
import org.celper.DefaultValue;
import org.celper.core.style._NoCellStyle;
import org.celper.processor.meta.FieldAnnotationMetaData;
import org.celper.processor.util.ElementUtil;

import javax.lang.model.element.VariableElement;
import javax.lang.model.type.TypeMirror;
import java.lang.annotation.Annotation;
import java.util.Objects;

public enum FieldAnnotationHandler implements AnnotationHandler<VariableElement, FieldAnnotationMetaData> {
    COLUMN(Column.class, (annotation, metaData, elementUtil) -> {
        Column column = (Column) annotation;
        metaData.setHeaderName(column.value());
        metaData.setPriority(column.priority());
    }),

    DEFAULT_VALUE(DefaultValue.class, (annotation, metaData, elementUtil) -> {
        DefaultValue defaultValue = (DefaultValue) annotation;
        metaData.setDefaultValue(defaultValue.value());
    }),

    CELL_FORMAT(CellFormat.class, (annotation, metaData, elementUtil) -> {
        CellFormat cellFormat = (CellFormat) annotation;
        metaData.setCellFormat(!"".equals(cellFormat.customFormat()) ? cellFormat.customFormat() : cellFormat.builtinFormat().getCellFormat());
    }),

    COLUMN_STYLE(ColumnStyle.class, (annotation, metaData, elementUtil) -> {
        ColumnStyle columnStyle = (ColumnStyle) annotation;
        TypeMirror headerStyleTypeMirror = elementUtil.getTypeMirror(columnStyle :: headerAreaStyle);
        TypeMirror dataStyleTypeMirror = elementUtil.getTypeMirror(columnStyle :: dataAreaStyle);
        headerStyleTypeMirror = elementUtil.isSameType(headerStyleTypeMirror, _NoCellStyle.class) ? null : headerStyleTypeMirror;
        dataStyleTypeMirror = elementUtil.isSameType(dataStyleTypeMirror, _NoCellStyle.class) ? null : dataStyleTypeMirror;
        metaData.setHeaderStyleConfigurerTypeMirror(headerStyleTypeMirror);
        metaData.setDataStyleConfigurerTypeMirror(dataStyleTypeMirror);
    });

    private final Class<? extends Annotation> type;
    private final TriConsumer<Annotation, FieldAnnotationMetaData, ElementUtil> ifPresent;

    FieldAnnotationHandler(Class<? extends Annotation> type,
                           TriConsumer<Annotation, FieldAnnotationMetaData, ElementUtil> ifPresent) {
        this.type = type;
        this.ifPresent = ifPresent;
    }

    @Override
    public void ifPresent(VariableElement field,
                          FieldAnnotationMetaData metaData,
                          ElementUtil elementUtil) {
        Annotation annotation = field.getAnnotation(this.type);
        if (Objects.nonNull(annotation))
            this.ifPresent.accept(annotation, metaData, elementUtil);
    }

}
