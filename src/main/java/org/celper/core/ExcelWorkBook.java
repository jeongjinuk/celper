package org.celper.core;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.celper.exception.EmptySheetException;
import org.celper.type.WorkBookType;
import org.celper.util.StructureRegistrator;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

public class ExcelWorkBook {
    private final Workbook _wb;
    private final List<ExcelSheet> sheets = new ArrayList<>();
    private final StructureRegistrator structureRegistrator;

    public ExcelWorkBook(WorkBookType workBookType) {
        this(workBookType.createWorkBook());
    }
    public ExcelWorkBook(WorkBookType workBookType, StructureRegistrator structureRegistrator){
        this(workBookType.createWorkBook(), structureRegistrator);
    }
    public ExcelWorkBook(Workbook workbook){
        this(workbook, new StructureRegistrator());
    }
    public ExcelWorkBook(Workbook workbook, StructureRegistrator structureRegistrator) {
        this._wb = workbook;
        this.structureRegistrator = structureRegistrator;
        workbook.spliterator().forEachRemaining(sheet -> this.sheets.add(new ExcelSheet(workbook, sheet, structureRegistrator)));
    }

    public ExcelSheet createSheet() {
        ExcelSheet sheet = new ExcelSheet(this._wb, this._wb.createSheet(), this.structureRegistrator);
        this.sheets.add(sheet);
        return sheet;
    }

    public ExcelSheet createSheet(String name) {
        ExcelSheet sheet = new ExcelSheet(this._wb, this._wb.createSheet(name), this.structureRegistrator);
        this.sheets.add(sheet);
        return sheet;
    }

    public ExcelSheet getSheetAt(int idx) {
        return this.sheets.get(idx);
    }

    public ExcelSheet getSheetByName(String name) {
        return this.sheets.stream()
                .filter(excelSheet -> name.equals(excelSheet.getName()))
                .findAny()
                .orElseThrow(() -> new EmptySheetException(String.format("%s 의 시트는 존재하지 않습니다.", name)));
    }

    public int size() {
        return this.sheets.size();
    }

    public Workbook getWorkbook() {
        return _wb;
    }

    public void write(OutputStream outputStream) throws IOException {
        this._wb.write(outputStream);
        _wb.close();
        if (_wb.getClass().equals(SXSSFWorkbook.class)){
            ((SXSSFWorkbook)_wb).dispose();
        }
        outputStream.close();
    }
}
