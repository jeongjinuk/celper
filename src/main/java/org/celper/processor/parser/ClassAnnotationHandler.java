package org.celper.processor.parser;

import org.celper.SheetStyle;
import org.celper.core.style.SheetStyleConfigurer;
import org.celper.processor.meta.ClassMetaData;

import javax.lang.model.element.Element;
import java.lang.annotation.Annotation;
import java.util.Objects;
import java.util.function.BiConsumer;

public enum ClassAnnotationHandler implements AnnotationHandler<Element, ClassMetaData>{
    COLUMN(SheetStyle.class, (annotation, classMetaData) -> {
        SheetStyle sheetStyle = (SheetStyle) annotation;
        classMetaData.setSheetStyleConfigurer((Class<SheetStyleConfigurer>) sheetStyle.value());
    });

    private final Class<? extends Annotation> type;
    private final BiConsumer<Annotation, ClassMetaData> ifPresent;

    ClassAnnotationHandler(Class<? extends Annotation> type, BiConsumer<Annotation, ClassMetaData> ifPresent) {
        this.type = type;
        this.ifPresent = ifPresent;
    }
    @Override
    public void ifPresent(Element field,
                         ClassMetaData metaData) {
        Annotation annotation = field.getAnnotation(this.type);
        if (Objects.nonNull(annotation))
            this.ifPresent.accept(annotation, metaData);
    }
}
