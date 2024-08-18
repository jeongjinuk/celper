package org.celper.core.writer;

import lombok.Getter;
import org.celper.core.common.CSVQuoteStrategy;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.Closeable;
import java.io.IOException;
import java.io.UncheckedIOException;
import java.io.Writer;
import java.nio.CharBuffer;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.StreamSupport;


public class CSVWriter<T> implements Closeable {
    Logger log = LoggerFactory.getLogger(CSVWriter.class);
    private final Writer writer;
    private final List<Function<T, Object>> getters; // T 타입의 필드 가져오는 방법
    private final List<String> columnNames;
    private final List<String> defaultValues;
    private final String lineDelimiter; // 줄바꿈
    private final String fieldSeparator; // 필드 나누기 문자열
    private final char quote; // 인용구 "" 감싸기
    private final String commentPrefix; // 주석 추가
    private final CSVQuoteStrategy[] csvQuoteStrategy;

    private int fieldLength;

    CSVWriter(WriterBuilder<T> builder) {
        this.writer = builder.getWriter();
        this.getters = builder.getGetters();
        this.defaultValues = builder.getDefaultValues();
        this.columnNames = builder.getColumnNames();

        this.lineDelimiter = builder.getLineDelimiter();
        this.fieldSeparator = builder.getFieldSeparator();
        this.quote = builder.getQuote();
        this.commentPrefix = builder.getCommentPrefix();

        this.fieldLength = builder.getFieldLength();
        this.csvQuoteStrategy = builder.getCsvQuoteStrategies();

    }
    public void writeRecords(Iterable<T> values) {
        try {
            writer.write(joining(columnNames.iterator()));
            for (T record : values) {
                writeRecord(record);
            }
            writer.flush();
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }

    }
    public void writeRecord(String defaultValue, String... values){
        try {
            defaultValue = Objects.toString(defaultValue, "");
            for (int i = 0; i < values.length; i++) {
                writeField(i, values[i], defaultValue);
            }
            writer.write(lineDelimiter);
            writer.flush();
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }
    public void writeComment(String comment) {
        try {
            writer.write(commentPrefix + comment + lineDelimiter);
        } catch (IOException e) {
            throw new UncheckedIOException(e);
        }
    }
    public void writeRecord(T t) {
        try{
            for (int i = 0; i < fieldLength; i++) {
                writeField(i, t);
            }
            writer.write(lineDelimiter);
        }catch (IOException e){
            throw new UncheckedIOException(e);
        }
    }

    private void writeField(int idx, String value, String defaultValue) throws IOException{
        if (idx > 0 && fieldLength > 1) writer.write(fieldSeparator);
        String field = Objects.toString(value, defaultValue);

        if (!requiresQuotes(field)) {
            writer.write(field);
        } else {
            writer.write(quote);
            appendEscape(field);
            writer.write(quote);
        }
    }
    private void writeField(int idx, T t) throws IOException {
        if (idx > 0 && fieldLength > 1) writer.write(fieldSeparator);

        String field = Objects.toString(getters.get(idx).apply(t), defaultValues.get(idx));

        if (!requiresQuotes(field)) {
            writer.write(field);
        } else {
            writer.write(quote);
            appendEscape(field);
            writer.write(quote);
        }
    }
    private void appendEscape(String field) throws IOException {
        int startPos = 0;
        int nextQuotePos = field.indexOf(quote, startPos);
        while (nextQuotePos != -1) {
            writer.write(field, startPos, nextQuotePos + 1);
            writer.write(quote);
            startPos = nextQuotePos + 1;
            nextQuotePos = field.indexOf(quote, startPos);
        }
        writer.write(field, startPos, field.length());
    }
    private boolean requiresQuotes(String field) {
        if (field.contains(fieldSeparator) ||
                field.contains(lineDelimiter) ||
                field.contains(commentPrefix)){
            return true;
        }

        for (CSVQuoteStrategy csvQuoteStrategy : csvQuoteStrategy) {
            if (csvQuoteStrategy.contains(field)) return true;
        }
        return false;
    }
    private String joining(Iterator<String> iterator) {
        Function<Iterable<String>, Spliterator<String>> spliteratorFunction = Iterable :: spliterator;
        return StreamSupport.stream(spliteratorFunction.apply(() -> iterator), false)
                .collect(Collectors.joining(fieldSeparator, "", lineDelimiter));
    }

    @Override
    public void close() throws IOException {
        writer.close();
    }

    @Override
    public String toString() {
        return "CSVWriter{" +
                ", columnNames=" + columnNames +
                ", defaultValues=" + defaultValues +
                ", lineDelimiter='" + lineDelimiter + '\'' +
                ", fieldSeparator='" + fieldSeparator + '\'' +
                ", commentPrefix='" + commentPrefix + '\'' +
                ", fieldLength=" + fieldLength +
                '}';
    }

    private static class InternalWriter extends Writer {

        private final Writer writer;
        private final CharBuffer charBuffer;
        private final int bufferCapacity;
        private int pos;

        InternalWriter(Writer writer, CharBuffer charBuffer, int bufferCapacity) {
            this.writer = writer;
            this.charBuffer = charBuffer;
            this.bufferCapacity = bufferCapacity;
        }

        @Override
        public void write(int c) throws IOException {
            if (pos >= bufferCapacity) {
                flush();
            }
            charBuffer.put((char) c);
            pos++;
        }

        @Override
        public void write(String str, int off, int len) throws IOException {
            while (off < len) {
                int remaining = bufferCapacity - pos;
                int length = Math.min(remaining, len - off);
                charBuffer.put(str, off, off + length);
                pos += length;
                off += length;
                if (pos >= bufferCapacity) {
                    flush();
                }
            }
        }

        @Override
        public void write(String str) throws IOException {
            write(str, 0, str.length());
        }

        @Override
        public void write(char[] cbuf, int off, int len) throws IOException {
            writer.write(cbuf, off, len);
        }

        @Override
        public void flush() throws IOException {
            if (pos > 0) {
                write(charBuffer.array(), 0, pos);
            }
            charBuffer.clear();
            pos = 0;
        }

        @Override
        public void close() throws IOException {
            flush();
            writer.close();
            charBuffer.clear();
        }

    }

    @Getter
    public static class WriterBuilder<T> {
        //TODO 네이밍 변경
        private Writer writer;
        private String lineDelimiter; // 줄바꿈
        private String fieldSeparator; // 필드 나누기 문자열
        private char quote; // 인용구 "" 감싸기
        private String commentPrefix; // 주석 추가
        private List<Function<T, Object>> getters; // T 타입의 필드 가져오는 방법
        /*
        기본 값 처리 없다면 ,[], <- []이것의 사이에는 아무것도 없음.
        만약 val1 = a, val2 = b, val3 = null, val4 = d 일경우
        val1,val2,val3,val4
        a,b,,d
        하지만 @DefaultValue("this is empty value")를 정의할 경우
        val1,val2,val3,val4
        a,b,this is empty value,d
        정의됨
         */
        private List<String> defaultValues;
        private List<String> columnNames;
        private int fieldLength;
        private int bufferCapacity;
        private CSVQuoteStrategy[] csvQuoteStrategies;

        private String[] skipHeaders;

        public WriterBuilder<T> setSkipHeaders(String[] skipHeaders) {
            this.skipHeaders = skipHeaders;
            return this;
        }

        public WriterBuilder<T> setBufferCapacity(int bufferCapacity) {
            this.bufferCapacity = bufferCapacity;
            return this;
        }

        public WriterBuilder<T> setLineDelimiter(String lineDelimiter) {
            this.lineDelimiter = lineDelimiter;
            return this;
        }

        public WriterBuilder<T> setFieldSeparator(String fieldSeparator) {
            this.fieldSeparator = fieldSeparator;
            return this;
        }

        public WriterBuilder<T> setQuoteStrategy(String quote) {
            this.quote = quote.charAt(0);
            return this;
        }

        public WriterBuilder<T> setCsvQuoteStrategies(CSVQuoteStrategy[] csvQuoteStrategies) {
            this.csvQuoteStrategies = csvQuoteStrategies;
            return this;
        }

        public WriterBuilder<T> setCommentPrefix(String commentPrefix) {
            this.commentPrefix = commentPrefix;
            return this;
        }

        public WriterBuilder<T> setGetters(List<Function<T, Object>> getters) {
            this.getters = getters;
            return this;
        }

        public WriterBuilder<T> setDefaultValues(List<String> defaultValues) {
            this.defaultValues = defaultValues;
            return this;
        }

        public WriterBuilder<T> setWriter(Writer writer) {
            this.writer = writer;
            return this;
        }

        public WriterBuilder<T> setColumnNames(List<String> columnNames) {
            this.columnNames = columnNames;
            return this;
        }

        private void setHeadersExcludingSkips(){
            if (Objects.isNull(skipHeaders)){
                return;
            }
            int[] nonSkipIndexes = IntStream.range(0, fieldLength)
                    .filter(i -> {
                        for (String skipHeader : skipHeaders) {
                            if (skipHeader.equals(columnNames.get(i))) return false;
                        }
                        return true;
                    })
                    .sorted()
                    .toArray();
            List<String> newHeaderNames = new ArrayList<>();
            List<Function<T, Object>> newGetters = new ArrayList<>();

            for (int i : nonSkipIndexes) {
                newHeaderNames.add(columnNames.get(i));
                newGetters.add(getters.get(i));
            }
            setColumnNames(newHeaderNames);
            setGetters(newGetters);
            this.fieldLength = getters.size();
        }

        public CSVWriter<T> build() {
            this.fieldLength = getters.size();
            setHeadersExcludingSkips();
            if (bufferCapacity > 0) {
                this.writer = new InternalWriter(writer, CharBuffer.allocate(bufferCapacity), bufferCapacity);
            }
            return new CSVWriter<>(this);
        }
    }
}
