package org.celper.core;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.celper.type.WorkBookType;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Optional;

/**
 * The type Excel work book.
 */
public class ExcelWorkBook {
    private final Workbook _wb;
    private final StructureRegistrator structureRegistrator;

    /**
     * Instantiates a new Excel work book.
     *
     * @param workBookType the work book type
     */
    public ExcelWorkBook(WorkBookType workBookType) {
        this(workBookType.createWorkBook());
    }

    /**
     * Instantiates a new Excel work book.
     *
     * @param workBookType         the work book type
     * @param structureRegistrator the structure registrator
     */
    public ExcelWorkBook(WorkBookType workBookType, StructureRegistrator structureRegistrator) {
        this(workBookType.createWorkBook(), structureRegistrator);
    }

    /**
     * Instantiates a new Excel work book.
     *
     * @param workbook the workbook
     */
    public ExcelWorkBook(Workbook workbook) {
        this(workbook, new StructureRegistrator());
    }

    /**
     * Instantiates a new Excel work book.
     *
     * @param workbook             the workbook
     * @param structureRegistrator the structure registrator
     */
    public ExcelWorkBook(Workbook workbook, StructureRegistrator structureRegistrator) {
        this._wb = workbook;
        this.structureRegistrator = structureRegistrator;
    }

    /**
     * Create sheet excel sheet.
     *
     * @return the excel sheet
     */
    public ExcelSheet createSheet() {
        ExcelSheet sheet = new ExcelSheet(this._wb, this._wb.createSheet(), this.structureRegistrator);
        return sheet;
    }

    /**
     * Create sheet excel sheet.
     *
     * @param name the name
     * @return the excel sheet
     */
    public ExcelSheet createSheet(String name) {
        ExcelSheet sheet = new ExcelSheet(this._wb, this._wb.createSheet(name), this.structureRegistrator);
        return sheet;
    }

    /**
     * Gets sheet at.
     *
     * @param idx the idx
     * @return the sheet at
     */
    public Optional<ExcelSheet> getSheetAt(int idx) {
        return Optional.ofNullable(new ExcelSheet(this._wb,this._wb.getSheetAt(idx), this.structureRegistrator));
    }

    /**
     * Gets sheet by name.
     *
     * @param name the name
     * @return the sheet by name
     */
    public Optional<ExcelSheet> getSheetByName(String name) {
        return Optional.ofNullable(new ExcelSheet(this._wb,this._wb.getSheet(name), this.structureRegistrator));
    }

    /**
     * Size int.
     *
     * @return the int
     */
    public int size() {
        return this._wb.getAllNames().size();
    }

    /**
     * Gets workbook.
     *
     * @return the workbook
     */
    public Workbook getWorkbook() {
        return _wb;
    }

    /**
     * Write.
     *
     * @param outputStream the output stream
     * @throws IOException the io exception
     */
    public void write(OutputStream outputStream) throws IOException {
        this._wb.write(outputStream);
        _wb.close();
        if (_wb.getClass().equals(SXSSFWorkbook.class)) {
            ((SXSSFWorkbook) _wb).dispose();
        }
        outputStream.close();
    }

}
