package org.celper.core.structure;

import lombok.*;
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

@Getter
@ToString
@Builder(access = AccessLevel.PRIVATE)
@AllArgsConstructor
public class ColumnStructure implements Comparable<ColumnStructure> {
    @Getter(AccessLevel.NONE)
    private Workbook _wb;
    private Structure structure;
    private List<String> importNameOptions;
    private CellStyle defaultCellStyle;
    private CellStyle headerAreaCellStyle;
    private CellStyle dataAreaCellStyle;
    private int headerRowPosition = -1;
    private int headerColumnPosition = -1;

    public ColumnStructure(Workbook workbook, Structure structure) {
        this._wb = workbook;
        this.structure = structure;
        this.importNameOptions = createImportNameOptions();
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

    public ColumnStructure newColumnStructure(int row, int col) {
        return ColumnStructure.builder()
                ._wb(this._wb)
                .structure(this.structure)
                .importNameOptions(this.importNameOptions)
                .defaultCellStyle(this.defaultCellStyle)
                .headerAreaCellStyle(this.headerAreaCellStyle)
                .dataAreaCellStyle(this.dataAreaCellStyle)
                .headerRowPosition(row)
                .headerColumnPosition(col)
                .build();
    }

    public int compareRowPosition(ColumnStructure o) {
        return o.getHeaderRowPosition() - this.getHeaderRowPosition();
    }

    public boolean isDefaultValueExists() {
        return !"".equals(this.structure.getDefaultValue());
    }

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
