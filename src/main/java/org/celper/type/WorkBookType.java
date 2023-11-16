package org.celper.type;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.celper.exception.UnsupportedWorkBookVersionException;

/**
 * The enum Work book type.
 */
public enum WorkBookType {
    /**
     * Hssf work book type.
     */
    HSSF("xls"),
    /**
     * Sxssf work book type.
     */
    SXSSF("xlsx"),
    /**
     * Xssf work book type.
     */
    XSSF("xlsx");
    private final String type;
    WorkBookType(String type) {
        this.type = type;
    }

    /**
     * Create work book workbook.
     *
     * @return the workbook
     */
    public Workbook createWorkBook() {
        switch (this){
            case HSSF:
                return new HSSFWorkbook();
            case SXSSF:
                return new SXSSFWorkbook();
            case XSSF:
                return new XSSFWorkbook();
            default:
                throw new UnsupportedWorkBookVersionException("Unsupported Workbook versions");
        }
    }
}

