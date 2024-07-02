package org.celper.processor;

import org.celper.core.style.SheetStyleConfigurer;

import java.util.List;

public class ClassMetaData {

    private Class<?> clazz;
    private SheetStyleConfigurer sheetStyleConfigurer;
    private List<FieldAnnotationMetaData> fieldAnnotationMetaDatas;

    public ClassMetaData() {}

    public ClassMetaData(Class<?> clazz,
                         SheetStyleConfigurer sheetStyleConfigurer,
                         List<FieldAnnotationMetaData> fieldAnnotationMetaDatas) {
        this.clazz = clazz;
        this.sheetStyleConfigurer = sheetStyleConfigurer;
        fieldAnnotationMetaDatas.sort(FieldAnnotationMetaData::compareTo);
        this.fieldAnnotationMetaDatas = fieldAnnotationMetaDatas;
    }

    public Class<?> getClazz() {
        return clazz;
    }

    public void setClazz(Class<?> clazz) {
        this.clazz = clazz;
    }

    public SheetStyleConfigurer getSheetStyleConfigurer() {
        return sheetStyleConfigurer;
    }

    public void setSheetStyleConfigurer(SheetStyleConfigurer sheetStyleConfigurer) {
        this.sheetStyleConfigurer = sheetStyleConfigurer;
    }

    public List<FieldAnnotationMetaData> getFieldAnnotationMetaDatas() {
        return fieldAnnotationMetaDatas;
    }

    public void setFieldAnnotationMetaDatas(List<FieldAnnotationMetaData> fieldAnnotationMetaDatas) {
        this.fieldAnnotationMetaDatas = fieldAnnotationMetaDatas;
    }
}
