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

public class ColumnStructure implements Comparable<ColumnStructure> {
    private Workbook _wb;
    private Structure structure;
    private List<String> importNameOptions;
    private CellStyle defaultCellStyle;
    private CellStyle headerAreaCellStyle;
    private CellStyle dataAreaCellStyle;
    public ColumnStructure(Workbook workbook, Structure structure) {
        this._wb = workbook;
        this.structure = structure;
        this.importNameOptions = createImportNameOptions();
    }

    public Workbook get_wb() {
        return _wb;
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

    public void setColumnStyle() {
        this.headerAreaCellStyle = createCellStyle(structure :: getHeaderAreaConfigurer);
        this.dataAreaCellStyle = createCellStyle(structure :: getDataAreaConfigurer);
        setDataFormat();
    }

    public void setSheetStyle(Sheet sheet) {
        this.defaultCellStyle = this._wb.createCellStyle();
        SheetStyleBuilder builder = new SheetStyleBuilder(this._wb, sheet, this.defaultCellStyle);
        this.structure.getSheetStyleConfigurer().config(builder);
    }

    public void setNonSheetStyle() {
        this.defaultCellStyle = this._wb.createCellStyle();
    }

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
