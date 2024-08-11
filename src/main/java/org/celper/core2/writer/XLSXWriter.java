package org.celper.core2.writer;

import lombok.Getter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.celper.MetaData;
import org.celper.core2.Util;
import org.celper.core2.style.CellStyleConfigurer;
import org.celper.core2.style.builder.CellStyleBuilder;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

public class XLSXWriter<T>{

    private static final Logger log = LoggerFactory.getLogger(XLSXWriter.class);
    private final SXSSFWorkbook workBook;
    private final MetaData<T> metaData;
    private final List<CellStyle> headerStyles;
    private final List<CellStyle> dataStyles;
    private final int colCount;

    private XLSXWriter(WriterBuilder<T> builder) {
        this.workBook = builder.getWorkBook();
        this.metaData = builder.getMetaData();
        this.headerStyles = builder.getHeaderStyles();
        this.dataStyles = builder.getDataStyles();
        this.colCount = builder.getColCount();
    }

    public SXSSFWorkbook getWorkBook() {
        return workBook;
    }

    public XLSXWriter<T> addSheet(List<T> list) {
        addSheet(list, null);
        return this;
    }

    public XLSXWriter<T> addSheet(List<T> list, String name) {
        workBook.setMissingCellPolicy(Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        SXSSFSheet sheet = Objects.isNull(name) ? this.workBook.createSheet() : this.workBook.createSheet(name);
        int size = list.size();
        setCellValue(sheet.createRow(0), metaData.getColumnNameList(), headerStyles);

        for (int rowNum = 1; rowNum < size; rowNum++) {
            setCellValue(sheet.createRow(rowNum), getValues(list.get(rowNum -1)), dataStyles);
        }
        return this;
    }


    private List<Object> getValues(T t){
        return metaData.getGetterFunctionList().stream()
                .map(f -> f.apply(t))
                .collect(Collectors.toList());
    }


    private void setCellValue(Row row, List<? extends Object> objects, List<CellStyle> styles){
        for (int col = 0; col < colCount; col++) {
            Object val = objects.get(col);
            Cell cell = row.getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            if (Objects.nonNull(val)){
                Util.setValue(cell, val);
            }

            if (existDefaultValue(val, col)){
                Util.setValue(cell, val);
            }
            cell.setCellStyle(styles.get(col));
        }
    }

    private boolean existDefaultValue(Object val, int col){
        return Objects.isNull(val) && Objects.nonNull(metaData.getDefaultValueList().get(col));
    }

    public void write(OutputStream stream) throws IOException {
        workBook.write(stream);
        workBook.dispose();
        stream.close();
    }


    @Getter
    public static class WriterBuilder<T>{
        private SXSSFWorkbook workBook;
        private List<CellStyle> headerStyles;
        private List<CellStyle> dataStyles;
        private MetaData<T> metaData;
        private int colCount;

        public WriterBuilder<T> setMetaData(MetaData<T> metaData) {
            this.metaData = metaData;
            return this;
        }

        public WriterBuilder<T> setWorkBook(SXSSFWorkbook workBook) {
            this.workBook = workBook;
            return this;
        }

        public int getColCount() {
            return metaData.getGetterFunctionList().size();
        }

        private CellStyle getDefaultStyle(){
            CellStyle defaultStyle = workBook.createCellStyle();
            metaData.getSheetStyle().config(new CellStyleBuilder(defaultStyle, workBook.createFont()));
            return defaultStyle;
        }

        private List<CellStyle> makeStyles(CellStyle defaultStyle, List<CellStyleConfigurer> list){
            return list.stream()
                    .map(cellStyleConfigurer -> makeStyle(cellStyleConfigurer, defaultStyle))
                    .collect(Collectors.toList());
        }

        private CellStyle makeStyle(CellStyleConfigurer cellStyleConfigurer, CellStyle defaultStyle) {
            if (Objects.isNull(cellStyleConfigurer)) {
                return defaultStyle;
            }
            CellStyle cellStyle = this.workBook.createCellStyle();
            Font font = this.workBook.createFont();

            cellStyle.cloneStyleFrom(defaultStyle);
            cellStyleConfigurer.config(new CellStyleBuilder(cellStyle, font));
            return cellStyle;
        }

        private void setFormats(List<CellStyle> list){
            List<String> cellFormatList = metaData.getCellFormatList();
            IntStream.rangeClosed(0, colCount)
                    .filter(i -> Objects.nonNull(cellFormatList.get(i)))
                    .forEach(i -> list.get(i).setDataFormat(workBook.createDataFormat().getFormat(cellFormatList.get(i))));
        }

        public XLSXWriter<T> build(){
            CellStyle defaultStyle = getDefaultStyle();
            this.headerStyles = makeStyles(defaultStyle, metaData.getHeaderStyleConfigList());
            this.dataStyles = makeStyles(defaultStyle, metaData.getDataStyleConfigList());
            setFormats(dataStyles);
            this.colCount = getColCount();
            return new XLSXWriter<>(this);
        }
    }
}
