package org.celper.processor.meta;

import org.celper.core.style.SheetStyleConfigurer;

import javax.lang.model.element.Element;
import javax.lang.model.element.TypeElement;
import java.util.List;
import java.util.function.Function;
import java.util.stream.Collectors;

public class ClassMetaData {

    private TypeElement clazz;
    private Class<SheetStyleConfigurer> sheetStyleConfigurer;
    private List<FieldAnnotationMetaData> fieldAnnotationMetaDataList;

    public ClassMetaData(TypeElement clazz, List<FieldAnnotationMetaData> fieldAnnotationMetaDataList) {
        this.clazz = clazz;
        this.fieldAnnotationMetaDataList = fieldAnnotationMetaDataList;
    }

    public <R> List<R> getFieldList(Function<FieldAnnotationMetaData, R> function){
        return fieldAnnotationMetaDataList.stream()
                .map(function::apply)
                .collect(Collectors.toList());
    }


    public TypeElement getClazz() {
        return clazz;
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
