package org.celper.processor.meta;

import lombok.Getter;
import lombok.Setter;

import javax.lang.model.element.VariableElement;
import javax.lang.model.type.TypeMirror;

@Getter
@Setter
public class FieldAnnotationMetaData implements Comparable<FieldAnnotationMetaData>{
    private VariableElement field;
    private String getterMethod;
    private int priority;
    private int definedFieldOrder;
    private String headerName;

    private String cellFormat;

    private String defaultValue;

    private TypeMirror headerStyleConfigurerTypeMirror;
    private TypeMirror dataStyleConfigurerTypeMirror;

    // construct
    public FieldAnnotationMetaData() {}

    // getters and setters

    @Override
    public int compareTo(FieldAnnotationMetaData o) {
        if (this.priority == 0 && o.getPriority() == 0)
            return Integer.compare(this.definedFieldOrder, o.getDefinedFieldOrder());

        if (this.priority == o.getPriority())
            return this.headerName.compareTo(o.getHeaderName());

        return Integer.compare(o.getPriority(), this.priority);
    }
}
