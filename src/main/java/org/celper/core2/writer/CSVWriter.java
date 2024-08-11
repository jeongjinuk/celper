package org.celper.core2.writer;

import lombok.Getter;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.io.Writer;
import java.nio.CharBuffer;
import java.util.*;
import java.util.function.Function;
import java.util.function.IntFunction;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * 1. Writer 받아서 사용하는 형태가 좋을 듯 등록해버리면 안되니깐
 * 2. line 별로 코멘트 넣기
 * 3.
 */
public class CSVWriter<T> {
    private final Writer writer;
    private final List<Function<T, Object>> getters; // T 타입의 필드 가져오는 방법
    private final List<String> columnNames;
    private final List<String> defaultValues;
    private final CharSequence lineDelimiter; // 줄바꿈
    private final CharSequence fieldSeparator; // 필드 나누기 문자열
    private final CharSequence quote; // 인용구 "" 감싸기
    private final CharSequence commentPrefix; // 주석 추가
    private final int fieldLength;
    private final int charactorCount;
    private final IntFunction<Integer> bufferSizeFunction;
    private final int flushSize;
    private final List<String> cache = new ArrayList<>();
    private int capacity = 2048;
    private CharBuffer charBuffer = CharBuffer.allocate(capacity);

    private CSVWriter(WriterBuilder<T> builder) {
        this.writer = builder.getWriter();
        this.getters = builder.getGetters();
        this.defaultValues = builder.getDefaultValues();
        this.columnNames = builder.getColumnNames();

        this.lineDelimiter = builder.getLineDelimiter();
        this.fieldSeparator = builder.getFieldSeparator();
        this.quote = builder.getQuote();
        this.commentPrefix = builder.getCommentPrefix();

        this.fieldLength = builder.getFieldLength();
        this.charactorCount = (fieldLength * fieldSeparator.length()) - fieldSeparator.length() + lineDelimiter.length();
        this.flushSize = builder.getFlushSize();
        this.bufferSizeFunction = fieldLengthSum -> (fieldLengthSum * 4) + charactorCount;
    }
    enum CSVQuoteStrategy{
        DOUBLE_QUOTE("\""),
        LINE_FEED("\n"),
        CARRIAGE_RETURN("\r"),
        TAB("\t"),
        LINE_SEPARATOR("\u2028"),
        PARAGRAPH_SEPARATOR("\u2029"),
        NEXT_LINE("\u0085"),
        BACKSLASH("\\"),
        RECORD_SEPARATOR("\u001e"),
        UNIT_SEPARATOR("\u001f");
        private final CharSequence charSequence;

        CSVQuoteStrategy(CharSequence character) {
            this.charSequence = character;
        }

        public CharSequence getCharSequence() {
            return charSequence;
        }

        public static boolean contains(String str) {
            for (CSVQuoteStrategy strategy : values()) {
                if (str.contains(strategy.getCharSequence())) {
                    return true;
                }
            }
            return false;
        }
    }
    //TODO This로 다 바꿔야함
    public void writeRecords(List<T> t) {
        writeToOutput(joining(columnNames.iterator()));
        int size = t.size();
        for (int i = 0; i < size; i++) {
            write(t.get(i));
            flush(i);
        }
    }

    public void writeRecord(T t){
        write(t);
    }

    public void writeComment(String comment){
        writeToOutput(commentPrefix + comment);
    }
    private void write(T t) {
        int bufferSize = 0;
        String fieldChars = null;
        Object o;
        for (int i = 0; i < fieldLength; i++) {
            o = getters.get(i).apply(t);
            fieldChars = Objects.nonNull(o) ? o.toString() : defaultValues.get(i);
            bufferSize += fieldChars.length();
            cache.add(fieldChars);
        }

        if (capacity < (bufferSize = bufferSizeFunction.apply(bufferSize))){
            charBuffer = CharBuffer.allocate(capacity = bufferSize);
        }

        for (int i = 0; i < fieldLength; i++) {
            if (i > 0) charBuffer.append(fieldSeparator);
            put(cache.get(i));
        }
        charBuffer.append(lineDelimiter);
        charBuffer.flip();

        writeToOutput(charBuffer);

        charBuffer.clear();
        cache.clear();
    }

    private <O> void writeToOutput(O o){
        try {
            if (o instanceof String){
                writer.write((String) o);
            } else if (o instanceof CharBuffer) {
                writer.write(charBuffer.array(), charBuffer.position(), charBuffer.remaining());
            }
        }catch (IOException e){
            throw new UncheckedIOException(e);
        }
    }
    private void flush(int row){
        if (flushSize == 0 || row == 0) return;
        try{
            if (row % flushSize == 0) writer.flush();
        }catch (IOException e){
            throw new UncheckedIOException(e);
        }

    }
    private void put(String field) {
        if (!requiresQuotes(field)) {
            charBuffer.append(field);
            return;
        }
        charBuffer.append(quote);
        // TODO itar로 바꾼 이후 한번에 pos로 넣기
        for (char c : field.toCharArray()) {
            if (c == CSVQuoteStrategy.DOUBLE_QUOTE.charSequence.charAt(0)) {
                charBuffer.put(CSVQuoteStrategy.DOUBLE_QUOTE.charSequence.charAt(0)); // 이스케이프 처리
            }
            charBuffer.put(c);
        }
        charBuffer.append(quote);
    }

    private boolean requiresQuotes(String field){
        return CSVQuoteStrategy.contains(field) || field.contains(fieldSeparator) ||
                field.contains(lineDelimiter) || field.contains(commentPrefix);
    }
    private String joining(Iterator<String> iterator) {
        Function<Iterable<String>, Spliterator<String>> spliteratorFunction = Iterable :: spliterator;
        return StreamSupport.stream(spliteratorFunction.apply(() -> iterator),false)
                .collect(Collectors.joining(CharBuffer.wrap(fieldSeparator), "",  CharBuffer.wrap(lineDelimiter)));
    }

    @Getter
    public static class WriterBuilder<T> {

        //TODO 네이밍 변경
        private Writer writer;
        private CharSequence lineDelimiter; // 줄바꿈
        private CharSequence fieldSeparator; // 필드 나누기 문자열
        private CharSequence quote; // 인용구 "" 감싸기
        private CharSequence commentPrefix; // 주석 추가
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
        private int flushSize;

        public WriterBuilder<T> setFlushSize(int flushSize) {
            this.flushSize = flushSize;
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
            this.quote = quote;
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

        public CSVWriter<T> build() {
            this.fieldLength = getGetters().size();
            return new CSVWriter<>(this);
        }


    }
}
