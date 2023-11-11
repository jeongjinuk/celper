package org.celper.core.structure;

import org.celper.annotations.*;
import org.celper.core.style.CellStyleConfigurer;
import org.celper.core.style.SheetStyleConfigurer;
import org.celper.core.style._NoCellStyle;
import org.celper.type.BuiltinCellFormatType;
import org.celper.util.ReflectionUtils;

import java.lang.reflect.Field;

/**
 * The type Structure.
 */
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

    /**
     * Instantiates a new Structure.
     *
     * @param clazz        the clazz
     * @param field        the field
     * @param definedOrder the defined order
     */
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

    /**
     * Gets field.
     *
     * @return the field
     */
    public Field getField() {
        return field;
    }

    /**
     * Gets field name.
     *
     * @return the field name
     */
    public String getFieldName() {
        return fieldName;
    }

    /**
     * Gets column.
     *
     * @return the column
     */
    public Column getColumn() {
        return column;
    }

    /**
     * Gets default value.
     *
     * @return the default value
     */
    public String getDefaultValue() {
        return defaultValue;
    }

    /**
     * Gets cell format.
     *
     * @return the cell format
     */
    public String getCellFormat() {
        return cellFormat;
    }

    /**
     * Gets export priority.
     *
     * @return the export priority
     */
    public int getExportPriority() {
        return exportPriority;
    }

    /**
     * Gets sheet style configurer.
     *
     * @return the sheet style configurer
     */
    public SheetStyleConfigurer getSheetStyleConfigurer() {
        return sheetStyleConfigurer;
    }

    /**
     * Gets header area configurer.
     *
     * @return the header area configurer
     */
    public CellStyleConfigurer getHeaderAreaConfigurer() {
        return headerAreaConfigurer;
    }

    /**
     * Gets data area configurer.
     *
     * @return the data area configurer
     */
    public CellStyleConfigurer getDataAreaConfigurer() {
        return dataAreaConfigurer;
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


    /**
     * Exists column annotation boolean.
     *
     * @param field the field
     * @return the boolean
     */
    public static boolean existsColumnAnnotation(Field field){
        return field.isAnnotationPresent(Column.class);
    }
}
