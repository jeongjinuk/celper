package org.celper.core;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.celper.type.WorkBookType;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Optional;

public class ExcelWorkBook {
    private final Workbook _wb;
    private final StructureRegistrator structureRegistrator;

    public ExcelWorkBook(WorkBookType workBookType) {
        this(workBookType.createWorkBook());
    }

    public ExcelWorkBook(WorkBookType workBookType, StructureRegistrator structureRegistrator) {
        this(workBookType.createWorkBook(), structureRegistrator);
    }

    public ExcelWorkBook(Workbook workbook) {
        this(workbook, new StructureRegistrator());
    }

    public ExcelWorkBook(Workbook workbook, StructureRegistrator structureRegistrator) {
        this._wb = workbook;
        this.structureRegistrator = structureRegistrator;
    }

    public ExcelSheet createSheet() {
        ExcelSheet sheet = new ExcelSheet(this._wb, this._wb.createSheet(), this.structureRegistrator);
        return sheet;
    }

    public ExcelSheet createSheet(String name) {
        ExcelSheet sheet = new ExcelSheet(this._wb, this._wb.createSheet(name), this.structureRegistrator);
        return sheet;
    }

    public Optional<ExcelSheet> getSheetAt(int idx) {
        return Optional.ofNullable(new ExcelSheet(this._wb,this._wb.getSheetAt(idx), this.structureRegistrator));
    }

    public Optional<ExcelSheet> getSheetByName(String name) {
        return Optional.ofNullable(new ExcelSheet(this._wb,this._wb.getSheet(name), this.structureRegistrator));
    }

    public int size() {
        return this._wb.getAllNames().size();
    }

    public Workbook getWorkbook() {
        return _wb;
    }

    public void write(OutputStream outputStream) throws IOException {
        this._wb.write(outputStream);
        _wb.close();
        if (_wb.getClass().equals(SXSSFWorkbook.class)) {
            ((SXSSFWorkbook) _wb).dispose();
        }
        outputStream.close();
    }

}
