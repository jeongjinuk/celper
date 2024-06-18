package org.celper.core.structure;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public class ImportStructure implements Comparable<ImportStructure>{
    private Workbook _wb;
    private Structure structure;
    private List<String> importNameOptions;
    private CellStyle defaultCellStyle;
    private CellStyle headerAreaCellStyle;
    private CellStyle dataAreaCellStyle;
    private int headerRowPosition;
    private int headerColumnPosition;

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

    public Structure getStructure() {
        return structure;
    }

    public List<String> getImportNameOptions() {
        return importNameOptions;
    }

    public CellStyle getDefaultCellStyle() {
        return defaultCellStyle;
    }

    public CellStyle getHeaderAreaCellStyle() {
        return headerAreaCellStyle;
    }

    public CellStyle getDataAreaCellStyle() {
        return dataAreaCellStyle;
    }

    public int getHeaderRowPosition() {
        return headerRowPosition;
    }

    public int getHeaderColumnPosition() {
        return headerColumnPosition;
    }

    @Override
    public int compareTo(ImportStructure o) {
        return o.getHeaderRowPosition() - this.getHeaderRowPosition();
    }

    public boolean existPosition() {
        return headerRowPosition >= 0 && headerColumnPosition >= 0;
    }
}
