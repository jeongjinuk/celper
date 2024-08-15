package org.celper.core2.common;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.Calendar;
import java.util.Date;

public class Util {
    public static void setValue(Cell cell, Object o) {
        if (o instanceof Double) {
            cell.setCellValue((double) o);
        } else if (o instanceof Integer) {
            cell.setCellValue((int) o);
        } else if (o instanceof Long) {
            cell.setCellValue((long) o);
        } else if (o instanceof Boolean) {
            cell.setCellValue((Boolean) o);
        } else if (o instanceof RichTextString) {
            cell.setCellValue((RichTextString) o);
        } else if (o instanceof Date) {
            cell.setCellValue((Date) o);
        } else if (o instanceof Calendar) {
            cell.setCellValue((Calendar) o);
        } else if (o instanceof LocalDate) {
            cell.setCellValue(LocalDateTime.of((LocalDate) o, LocalTime.NOON));
        } else if (o instanceof LocalTime) {
            cell.setCellValue(LocalDateTime.of(LocalDate.MIN, (LocalTime) o));
        } else {
            cell.setCellValue(String.valueOf(o));
        }
    }
}
