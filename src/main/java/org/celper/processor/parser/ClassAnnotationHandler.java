package org.celper.processor.parser;

import org.celper.SheetStyle;
import org.celper.processor.meta.ClassMetaData;
import org.celper.processor.util.ElementUtil;

import javax.lang.model.element.TypeElement;
import java.lang.annotation.Annotation;
import java.util.Objects;
import java.util.function.BiConsumer;

public enum ClassAnnotationHandler implements AnnotationHandler<TypeElement, ClassMetaData>{
    COLUMN(SheetStyle.class, (annotation, classMetaData) -> {
        SheetStyle sheetStyle = (SheetStyle) annotation;
        classMetaData.setSheetStyleConfigurerTypeMirror(ElementUtil.getTypeMirror(sheetStyle::value));
    });

    private final Class<? extends Annotation> type;
    private final BiConsumer<Annotation, ClassMetaData> ifPresent;

    ClassAnnotationHandler(Class<? extends Annotation> type, BiConsumer<Annotation, ClassMetaData> ifPresent) {
        this.type = type;
        this.ifPresent = ifPresent;
    }

    @Override
    public void ifPresent(TypeElement clazz,
                         ClassMetaData metaData) {
        Annotation annotation = clazz.getAnnotation(this.type);
        if (Objects.nonNull(annotation))
            this.ifPresent.accept(annotation, metaData);
    }
}
