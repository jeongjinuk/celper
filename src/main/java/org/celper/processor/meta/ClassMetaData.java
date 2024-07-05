package org.celper.processor.meta;

import org.celper.core.style.SheetStyleConfigurer;

import javax.lang.model.element.Element;
import java.util.List;

public class ClassMetaData {

    private Element clazz;
    private Class<SheetStyleConfigurer> sheetStyleConfigurer;
    private List<FieldAnnotationMetaData> fieldAnnotationMetaDatas;

    public ClassMetaData(Element clazz, List<FieldAnnotationMetaData> fieldAnnotationMetaDatas) {
        this.clazz = clazz;
        this.fieldAnnotationMetaDatas = fieldAnnotationMetaDatas;
    }

    public Element getClazz() {
        return clazz;
    }

    public void setClazz(Element clazz) {
        this.clazz = clazz;
    }

    public Class<SheetStyleConfigurer> getSheetStyleConfigurer() {
        return sheetStyleConfigurer;
    }

    public void setSheetStyleConfigurer(Class<SheetStyleConfigurer> sheetStyleConfigurer) {
        this.sheetStyleConfigurer = sheetStyleConfigurer;
    }

    public List<FieldAnnotationMetaData> getFieldAnnotationMetaDatas() {
        return fieldAnnotationMetaDatas;
    }

    public void setFieldAnnotationMetaDatas(List<FieldAnnotationMetaData> fieldAnnotationMetaDatas) {
        this.fieldAnnotationMetaDatas = fieldAnnotationMetaDatas;
    }
}
