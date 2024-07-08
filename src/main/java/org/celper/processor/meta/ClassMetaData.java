package org.celper.processor.meta;

import org.celper.core.style.SheetStyleConfigurer;

import javax.lang.model.element.Element;
import javax.lang.model.element.TypeElement;
import java.util.List;

public class ClassMetaData {

    private TypeElement clazz;
    private Class<SheetStyleConfigurer> sheetStyleConfigurer;
    private List<FieldAnnotationMetaData> fieldAnnotationMetaDataList;

    public ClassMetaData(TypeElement clazz, List<FieldAnnotationMetaData> fieldAnnotationMetaDataList) {
        this.clazz = clazz;
        this.fieldAnnotationMetaDataList = fieldAnnotationMetaDataList;
    }

    public TypeElement getClazz() {
        return clazz;
    }

    public void setClazz(TypeElement clazz) {
        this.clazz = clazz;
    }

    public Class<SheetStyleConfigurer> getSheetStyleConfigurer() {
        return sheetStyleConfigurer;
    }

    public void setSheetStyleConfigurer(Class<SheetStyleConfigurer> sheetStyleConfigurer) {
        this.sheetStyleConfigurer = sheetStyleConfigurer;
    }

    public List<FieldAnnotationMetaData> getFieldAnnotationMetaDataList() {
        return fieldAnnotationMetaDataList;
    }

    public void setFieldAnnotationMetaDataList(List<FieldAnnotationMetaData> fieldAnnotationMetaDataList) {
        this.fieldAnnotationMetaDataList = fieldAnnotationMetaDataList;
    }
}
