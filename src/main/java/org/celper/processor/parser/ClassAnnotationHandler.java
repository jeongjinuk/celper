package org.celper.processor.parser;

import org.celper.SheetStyle;
import org.celper.core.style._NoCellStyle;
import org.celper.processor.meta.ClassMetaData;
import org.celper.processor.util.ElementUtil;

import javax.lang.model.element.TypeElement;
import javax.lang.model.type.TypeMirror;
import java.lang.annotation.Annotation;
import java.util.Objects;

public enum ClassAnnotationHandler implements AnnotationHandler<TypeElement, ClassMetaData>{
    COLUMN(SheetStyle.class, (annotation, classMetaData, elementUtil) -> {
        SheetStyle sheetStyle = (SheetStyle) annotation;
        TypeMirror typeMirror = elementUtil.getTypeMirror(sheetStyle :: value);
        typeMirror = elementUtil.isSameType(typeMirror, _NoCellStyle.class) ? null : typeMirror;
        classMetaData.setSheetStyleConfigurerTypeMirror(typeMirror);
    });

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
