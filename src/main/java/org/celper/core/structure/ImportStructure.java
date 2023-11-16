package org.celper.core.structure;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

/**
 * The type Import structure.
 */
public class ImportStructure implements Comparable<ImportStructure>{
    private Workbook _wb;
    private Structure structure;
    private List<String> importNameOptions;
    private CellStyle defaultCellStyle;
    private CellStyle headerAreaCellStyle;
    private CellStyle dataAreaCellStyle;
    private int headerRowPosition;
    private int headerColumnPosition;

    /**
     * Instantiates a new Import structure.
     *
     * @param columnStructure      the column structure
     * @param headerRowPosition    the header row position
     * @param headerColumnPosition the header column position
     */
    public ImportStructure(ColumnStructure columnStructure,
                            int headerRowPosition,
                            int headerColumnPosition) {
        this._wb = columnStructure.get_wb();
        this.structure = columnStructure.getStructure();
        this.importNameOptions = columnStructure.getImportNameOptions();
        this.defaultCellStyle = columnStructure.getDefaultCellStyle();
        this.headerAreaCellStyle = columnStructure.getHeaderAreaCellStyle();
        this.dataAreaCellStyle = columnStructure.getDataAreaCellStyle();
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

    @Override
    public int compareTo(ImportStructure o) {
        return o.getHeaderRowPosition() - this.getHeaderRowPosition();
    }

    /**
     * Exist position boolean.
     *
     * @return the boolean
     */
    public boolean existPosition() {
        return headerRowPosition >= 0 && headerColumnPosition >= 0;
    }
}
