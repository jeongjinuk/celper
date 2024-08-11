package org.celper;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.celper.core2.writer.CSVWriter;
import org.celper.core2.writer.XLSXWriter;
import org.celper.register.RegisterManager;

import java.io.Writer;
import java.util.function.Function;

public class CelperFactory {

    private static final Function<Class<?>, MetaData<?>> META_DATA_FUNCTION = RegisterManager :: getMetaData;
    private static int DEFAULT_FLUSH_SIZE = 100;
    private static boolean DEFAULT_COMPRESS_TMP_FILE = false;
    private static boolean USE_SHARED_STRINGS_TABLE = false;

    private CelperFactory() {}

    public static <T> XLSXWriter<T> xlsxWriter(Class<T> clazz) {
        return xlsxWriter(clazz, DEFAULT_FLUSH_SIZE, DEFAULT_COMPRESS_TMP_FILE, USE_SHARED_STRINGS_TABLE);
    }
    public static <T> XLSXWriter<T> xlsxWriter(Class<T> clazz,
                                               int flushSize) {
        return xlsxWriter(clazz, flushSize, DEFAULT_COMPRESS_TMP_FILE, USE_SHARED_STRINGS_TABLE);
    }
    public static <T> XLSXWriter<T> xlsxWriter(Class<T> clazz,
                                               int flushSize,
                                               boolean compressTmpFiles,
                                               boolean useSharedStringsTable) {
        return new XLSXWriter.WriterBuilder<T>()
                .setWorkBook(new SXSSFWorkbook(new XSSFWorkbook(), flushSize, compressTmpFiles, useSharedStringsTable))
                .setMetaData((MetaData<T>) META_DATA_FUNCTION.apply(clazz))
                .build();
    }

    public static <T> CSVWriter<T> csvWriter(Class<T> clazz, Writer writer, int flushRow) {
        MetaData<T> metaData = (MetaData<T>) META_DATA_FUNCTION.apply(clazz);
        String[] csvConfig = metaData.getCsvConfig();
        return new CSVWriter.WriterBuilder<T>()
                .setWriter(writer)
                .setGetters(metaData.getGetterFunctionList())
                .setColumnNames(metaData.getColumnNameList())
                .setDefaultValues(metaData.getDefaultValueList())
                .setLineDelimiter(csvConfig[0])
                .setFieldSeparator(csvConfig[1])
                .setQuoteStrategy(csvConfig[2])
                .setCommentPrefix(csvConfig[3])
                .setFlushSize(flushRow)
                .build();
    }


    public static <T> CSVWriter<T> csvWriter(Class<T> clazz, Writer writer) {
        MetaData<T> metaData = (MetaData<T>) META_DATA_FUNCTION.apply(clazz);
        String[] csvConfig = metaData.getCsvConfig();
        return new CSVWriter.WriterBuilder<T>()
                .setWriter(writer)
                .setGetters(metaData.getGetterFunctionList())
                .setColumnNames(metaData.getColumnNameList())
                .setDefaultValues(metaData.getDefaultValueList())
                .setLineDelimiter(csvConfig[0])
                .setFieldSeparator(csvConfig[1])
                .setQuoteStrategy(csvConfig[2])
                .setCommentPrefix(csvConfig[3])
                .build();
    }


    // CSVWriter 추가





}
