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
    private int headerRowPosition = -1;
    private int headerColumnPosition = -1;

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

    private ColumnStructure(Workbook _wb,
                            Structure structure,
                            List<String> importNameOptions,
                            CellStyle defaultCellStyle,
                            CellStyle headerAreaCellStyle,
                            CellStyle dataAreaCellStyle,
                            int headerRowPosition,
                            int headerColumnPosition) {
        this._wb = _wb;
        this.structure = structure;
        this.importNameOptions = importNameOptions;
        this.defaultCellStyle = defaultCellStyle;
        this.headerAreaCellStyle = headerAreaCellStyle;
        this.dataAreaCellStyle = dataAreaCellStyle;
        this.headerRowPosition = headerRowPosition;
        this.headerColumnPosition = headerColumnPosition;
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
     * Gets header row position.
     *
     * @return the header row position
     */
    public int getHeaderRowPosition() {
        return headerRowPosition;
    }

    /**
     * Gets header column position.
     *
     * @return the header column position
     */
    public int getHeaderColumnPosition() {
        return headerColumnPosition;
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
     * New column structure column structure.
     *
     * @param row the row
     * @param col the col
     * @return the column structure
     */
    public ColumnStructure newColumnStructure(int row, int col) {
        return new ColumnStructure(
                this._wb,
                this.structure,
                this.importNameOptions,
                this.defaultCellStyle,
                this.headerAreaCellStyle,
                this.dataAreaCellStyle,
                row,
                col);
    }

    /**
     * Compare row position int.
     *
     * @param o the o
     * @return the int
     */
    public int compareRowPosition(ColumnStructure o) {
        return o.getHeaderRowPosition() - this.getHeaderRowPosition();
    }

    /**
     * Is default value exists boolean.
     *
     * @return the boolean
     */
    public boolean isDefaultValueExists() {
        return !"".equals(this.structure.getDefaultValue());
    }

    /**
     * Exist position boolean.
     *
     * @return the boolean
     */
    public boolean existPosition() {
        return headerRowPosition >= 0 && headerColumnPosition >= 0;
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
