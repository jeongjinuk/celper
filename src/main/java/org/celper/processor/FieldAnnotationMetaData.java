package org.celper.processor;

import org.celper.core.style.CellStyleConfigurer;

import javax.lang.model.element.ExecutableElement;
import javax.lang.model.element.VariableElement;

public class FieldAnnotationMetaData implements Comparable<FieldAnnotationMetaData>{
    private VariableElement field;
    private ExecutableElement method;
    private int priority;
    private int definedFieldOrder;
    private String headerName;
    private String CellFormat;
    private String defaultValue;
    private CellStyleConfigurer headerStyleConfigurer;
    private CellStyleConfigurer dataStyleConfigurer;

    // construct
    public FieldAnnotationMetaData() {}
    // getters and setters


    public VariableElement getField() {
        return field;
    }

    public void setField(VariableElement field) {
        this.field = field;
    }

    public ExecutableElement getMethod() {
        return method;
    }

    public void setMethod(ExecutableElement method) {
        this.method = method;
    }

    public int getPriority() {
        return priority;
    }

    public void setPriority(int priority) {
        this.priority = priority;
    }

    public int getDefinedFieldOrder() {
        return definedFieldOrder;
    }

    public void setDefinedFieldOrder(int definedFieldOrder) {
        this.definedFieldOrder = definedFieldOrder;
    }

    public String getHeaderName() {
        return headerName;
    }

    public void setHeaderName(String headerName) {
        this.headerName = headerName;
    }

    public String getCellFormat() {
        return CellFormat;
    }

    public void setCellFormat(String cellFormat) {
        CellFormat = cellFormat;
    }

    public String getDefaultValue() {
        return defaultValue;
    }

    public void setDefaultValue(String defaultValue) {
        this.defaultValue = defaultValue;
    }

    public CellStyleConfigurer getHeaderStyleConfigurer() {
        return headerStyleConfigurer;
    }

    public void setHeaderStyleConfigurer(CellStyleConfigurer headerStyleConfigurer) {
        this.headerStyleConfigurer = headerStyleConfigurer;
    }

    public CellStyleConfigurer getDataStyleConfigurer() {
        return dataStyleConfigurer;
    }

    public void setDataStyleConfigurer(CellStyleConfigurer dataStyleConfigurer) {
        this.dataStyleConfigurer = dataStyleConfigurer;
    }

    @Override
    public int compareTo(FieldAnnotationMetaData o) {
        if (this.priority == 0 && o.getPriority() == 0)
            return Integer.compare(this.definedFieldOrder, o.getDefinedFieldOrder());

        if (this.priority == o.getPriority())
            return this.headerName.compareTo(o.getHeaderName());

        return Integer.compare(o.getPriority(), this.priority);
    }
}
