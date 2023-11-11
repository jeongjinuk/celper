package org.celper.core.structure;

import org.celper.annotations.*;
import org.celper.core.style.CellStyleConfigurer;
import org.celper.core.style.SheetStyleConfigurer;
import org.celper.core.style._NoCellStyle;
import org.celper.type.BuiltinCellFormatType;
import org.celper.util.ReflectionUtils;

import java.lang.reflect.Field;

public class Structure {
    private final Field field;
    private String fieldName;
    private Column column;
    private String defaultValue;
    private String cellFormat;
    private int exportPriority;
    private SheetStyleConfigurer sheetStyleConfigurer = builder -> {};
    private CellStyleConfigurer headerAreaConfigurer = builder -> {};
    private CellStyleConfigurer dataAreaConfigurer = builder -> {};

    public  <T> Structure(final Class<T> clazz, final Field field, final int definedOrder) {
        this.field = field; // non null
        this.fieldName = field.getName();
        this.column = field.getDeclaredAnnotation(Column.class);
        setDefaultValue();
        setPriority(definedOrder);
        setSheetStyleConfigurer(clazz);
        setCellStyleConfigurer();
        setCellFormat();
    }

    private void setDefaultValue(){
        if (this.field.isAnnotationPresent(DefaultValue.class)){
            this.defaultValue = this.field.getDeclaredAnnotation(DefaultValue.class).value();
        }
    }

    private void setPriority(int definedOrder){
        this.exportPriority = this.column.priority() * 1000 + definedOrder;
    }

    private void setCellFormat() {
        if (!this.field.isAnnotationPresent(CellFormat.class)){
            this.cellFormat = BuiltinCellFormatType.GENERAL.getCellFormat();
            return;
        }
        CellFormat annotation = this.field.getDeclaredAnnotation(CellFormat.class);
        this.cellFormat = "".equals(annotation.customFormat()) ? annotation.builtinFormat().getCellFormat() : annotation.customFormat();
    }

    private void setSheetStyleConfigurer(Class<?> clazz) {
        if (clazz.isAnnotationPresent(SheetStyle.class)){
            this.sheetStyleConfigurer = ReflectionUtils.getInstance(clazz.getDeclaredAnnotation(SheetStyle.class).value());
        }
    }

    private void setCellStyleConfigurer(){
        if (this.field.isAnnotationPresent(ColumnStyle.class)){
            ColumnStyle annotation = this.field.getDeclaredAnnotation(ColumnStyle.class);
            if (!_NoCellStyle.class.equals(annotation.headerAreaStyle())){
                this.headerAreaConfigurer = ReflectionUtils.getInstance(annotation.headerAreaStyle());
            }
            if (!_NoCellStyle.class.equals(annotation.dataAreaStyle())){
                this.dataAreaConfigurer = ReflectionUtils.getInstance(annotation.dataAreaStyle());
            }
        }
    }


    public static boolean existsColumnAnnotation(Field field){
        return field.isAnnotationPresent(Column.class);
    }


    public Field getField() {
        return field;
    }
    public String getFieldName() {
        return fieldName;
    }
    public Column getColumn() {
        return column;
    }
    public String getDefaultValue() {
        return defaultValue;
    }
    public String getCellFormat() {
        return cellFormat;
    }
    public int getExportPriority() {
        return exportPriority;
    }
    public SheetStyleConfigurer getSheetStyleConfigurer() {
        return sheetStyleConfigurer;
    }
    public CellStyleConfigurer getHeaderAreaConfigurer() {
        return headerAreaConfigurer;
    }
    public CellStyleConfigurer getDataAreaConfigurer() {
        return dataAreaConfigurer;
    }
}
