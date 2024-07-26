package org.celper.processor.meta;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

import javax.lang.model.element.TypeElement;
import javax.lang.model.type.TypeMirror;
import java.util.List;
import java.util.function.Function;
import java.util.stream.Collectors;

@Getter
@Setter
@AllArgsConstructor
public class ClassMetaData {
    private TypeElement clazz;
    private TypeMirror sheetStyleConfigurerTypeMirror;
    private String[] csvConfig = new String[]{",", "\"\"", "\n"};
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

}
