package org.celper.core.structure;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.celper.core.style.CellStyleConfigurer;
import org.celper.core.style.builder.CellStyleBuilder;
import org.celper.core.style.builder.SheetStyleBuilder;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.function.Supplier;

/**
 * The type Column structure.
 */
public class ColumnStructure implements Comparable<ColumnStructure> {
    private Workbook _wb;
    private Structure structure;
    private List<String> importNameOptions;
    private CellStyle defaultCellStyle;
    private CellStyle headerAreaCellStyle;
    private CellStyle dataAreaCellStyle;

    /**
     * Instantiates a new Column structure.
     *
     * @param workbook  the workbook
     * @param structure the structure
     */
    public ColumnStructure(Workbook workbook, Structure structure) {
        this._wb = workbook;
        this.structure = structure;
        this.importNameOptions = createImportNameOptions();
    }

    /**
     * Gets wb.
     *
     * @return the wb
     */
    public Workbook get_wb() {
        return _wb;
    }

    /**
     * Gets structure.
     *
     * @return the structure
     */
    public Structure getStructure() {
        return structure;
    }

    /**
     * Gets import name options.
     *
     * @return the import name options
     */
    public List<String> getImportNameOptions() {
        return importNameOptions;
    }

    /**
     * Gets default cell style.
     *
     * @return the default cell style
     */
    public CellStyle getDefaultCellStyle() {
        return defaultCellStyle;
    }

    /**
     * Gets header area cell style.
     *
     * @return the header area cell style
     */
    public CellStyle getHeaderAreaCellStyle() {
        return headerAreaCellStyle;
    }

    /**
     * Gets data area cell style.
     *
     * @return the data area cell style
     */
    public CellStyle getDataAreaCellStyle() {
        return dataAreaCellStyle;
    }

    /**
     * Sets column style.
     */
    public void setColumnStyle() {
        this.headerAreaCellStyle = createCellStyle(structure :: getHeaderAreaConfigurer);
        this.dataAreaCellStyle = createCellStyle(structure :: getDataAreaConfigurer);
        setDataFormat();
    }

    /**
     * Sets sheet style.
     *
     * @param sheet the sheet
     */
    public void setSheetStyle(Sheet sheet) {
        this.defaultCellStyle = this._wb.createCellStyle();
        SheetStyleBuilder builder = new SheetStyleBuilder(this._wb, sheet, this.defaultCellStyle);
        this.structure.getSheetStyleConfigurer().config(builder);
    }

    /**
     * Sets non sheet style.
     */
    public void setNonSheetStyle() {
        this.defaultCellStyle = this._wb.createCellStyle();
    }

    /**
     * Is default value exists boolean.
     *
     * @return the boolean
     */
    public boolean isDefaultValueExists() {
        return !"".equals(this.structure.getDefaultValue());
    }

    private List<String> createImportNameOptions() {
        List<String> titles = new ArrayList<>();
        titles.add(this.structure.getColumn().value());
        if (!"".equals(this.structure.getColumn().importNameOptions()[0])) {
            titles.addAll(Arrays.asList(this.structure.getColumn().importNameOptions()));
        }
        return titles;
    }

    private CellStyle createCellStyle(Supplier<? extends CellStyleConfigurer> supplier) {
        CellStyle cellStyle = this._wb.createCellStyle();
        cellStyle.cloneStyleFrom(defaultCellStyle);
        supplier.get().config(new CellStyleBuilder(_wb, cellStyle));
        return cellStyle;
    }

    private void setDataFormat() {
        this.dataAreaCellStyle.setDataFormat(this._wb.createDataFormat().getFormat(this.structure.getCellFormat()));
    }

    @Override
    public int compareTo(ColumnStructure o) {
        boolean b = o.structure.getExportPriority() != this.structure.getExportPriority();
        return b ? this.structure.getExportPriority() - o.structure.getExportPriority() :
                this.structure.getColumn().value().compareTo(o.structure.getColumn().value());
    }
}
