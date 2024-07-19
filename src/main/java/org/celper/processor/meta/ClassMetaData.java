package org.celper.processor.meta;

import javax.lang.model.element.TypeElement;
import javax.lang.model.type.TypeMirror;
import java.util.List;
import java.util.function.Function;
import java.util.stream.Collectors;

public class ClassMetaData {

    private TypeElement clazz;
    private TypeMirror sheetStyleConfigurerTypeMirror;
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

    public TypeMirror getSheetStyleConfigurerTypeMirror() {
        return sheetStyleConfigurerTypeMirror;
    }

    public void setSheetStyleConfigurerTypeMirror(TypeMirror sheetStyleConfigurerTypeMirror) {
        this.sheetStyleConfigurerTypeMirror = sheetStyleConfigurerTypeMirror;
    }


    public List<FieldAnnotationMetaData> getFieldAnnotationMetaDataList() {
        return fieldAnnotationMetaDataList;
    }

    public void setFieldAnnotationMetaDataList(List<FieldAnnotationMetaData> fieldAnnotationMetaDataList) {
        this.fieldAnnotationMetaDataList = fieldAnnotationMetaDataList;
    }
}
