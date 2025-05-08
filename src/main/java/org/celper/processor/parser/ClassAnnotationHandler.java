package org.celper.processor.parser;

import org.celper.CSVConfig;
import org.celper.SheetLayout;
import org.celper.SheetStyle;
import org.celper.processor.meta.ClassMetaData;
import org.celper.processor.util.ElementUtil;

import javax.lang.model.element.TypeElement;
import javax.lang.model.type.TypeMirror;
import java.lang.annotation.Annotation;
import java.util.Objects;

public enum ClassAnnotationHandler implements AnnotationHandler<TypeElement, ClassMetaData>{
    SHEET_STYLE(SheetStyle.class, (annotation, classMetaData, elementUtil) -> {
        SheetStyle sheetStyle = (SheetStyle) annotation;
        TypeMirror typeMirror = elementUtil.getTypeMirror(sheetStyle :: value);
        classMetaData.setSheetStyleConfigurerTypeMirror(typeMirror);
    }),
    SHEET_LAYOUT_CONFIG(SheetLayout.class, (annotation, classMetaData, elementUtil) -> {
        SheetLayout sheetLayout = (SheetLayout) annotation;
        TypeMirror typeMirror = elementUtil.getTypeMirror(sheetLayout  :: value);
        classMetaData.setSheetLayoutConfigurerTypeMirror(typeMirror);
    }),
    CSV_CONFIG(CSVConfig.class, (annotation, classMetaData, elementUtil) -> {
        CSVConfig csvModel = (CSVConfig) annotation;

        String lineDelimiter = !csvModel.lineDelimiter().isEmpty() ? csvModel.lineDelimiter() : CSVConfig.DefaultCSVConfig.LINE_DELIMITER_CR.getStr();
        String fieldSeparator = !csvModel.fieldSeparator().isEmpty() ? csvModel.fieldSeparator() : CSVConfig.DefaultCSVConfig.FIELD_SEPARATOR.getStr();
        String quoteStrategy = csvModel.quote() != '\0'? String.valueOf(csvModel.quote()) : CSVConfig.DefaultCSVConfig.QUOTE_STRATEGY.getStr();
        String commentPrefix = !csvModel.comment().isEmpty() ? csvModel.comment() :CSVConfig.DefaultCSVConfig.COMMENT_PREFIX.getStr();

        classMetaData.setCsvConfig(new String[]{lineDelimiter, fieldSeparator, quoteStrategy, commentPrefix});
    })
    ;

    private final Class<? extends Annotation> type;
    private final TriConsumer<Annotation, ClassMetaData, ElementUtil> ifPresent;

    ClassAnnotationHandler(Class<? extends Annotation> type,
                           TriConsumer<Annotation, ClassMetaData, ElementUtil> ifPresent) {
        this.type = type;
        this.ifPresent = ifPresent;
    }

    @Override
    public void ifPresent(TypeElement clazz,
                          ClassMetaData metaData,
                          ElementUtil elementUtil) {
        Annotation annotation = clazz.getAnnotation(this.type);
        if (Objects.nonNull(annotation))
            this.ifPresent.accept(annotation, metaData, elementUtil);
    }

}
