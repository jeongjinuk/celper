package org.celper;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.celper.core2.common.CSVQuoteStrategy;
import org.celper.core2.writer.*;
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

    public static <T> CSVWriter<T> csvWriter(Class<T> clazz, Writer writer) {
        return csvWriter(clazz, writer, 1024);
    }
    public static <T> CSVWriter<T> csvWriter(Class<T> clazz, Writer writer, String... skipHeaders) {
        return csvWriter(clazz, writer, 1024, skipHeaders);
    }
    public static <T> CSVWriter<T> csvWriter(Class<T> clazz, Writer writer, int bufferCapacity, String... skipHeaders) {
        MetaData<T> metaData = (MetaData<T>) META_DATA_FUNCTION.apply(clazz);
        String[] csvConfig = metaData.getCsvConfig();
        return new CSVWriter.WriterBuilder<T>()
                .setWriter(writer)
                .setGetters(metaData.getGetterFunctionList())
                .setColumnNames(metaData.getColumnNameList())
                .setDefaultValues(metaData.getDefaultValueList())
                .setCsvQuoteStrategies(CSVQuoteStrategy.values())
                .setLineDelimiter(csvConfig[0])
                .setFieldSeparator(csvConfig[1])
                .setQuoteStrategy(csvConfig[2])
                .setCommentPrefix(csvConfig[3])
                .setBufferCapacity(bufferCapacity)
                .setSkipHeaders(skipHeaders)
                .build();
    }





}
