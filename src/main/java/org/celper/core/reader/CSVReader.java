package org.celper.core.reader;

import org.celper.core.common.CSVQuoteStrategy;

import java.io.IOException;
import java.io.Reader;
import java.io.UncheckedIOException;
import java.nio.CharBuffer;
import java.util.ArrayDeque;
import java.util.function.Consumer;

/**
 * TODO
 * 1. Quote CRLF 추가해서 돌리게
 *
 */
public class CSVReader {

    private final InternalBuffer buffer;
    private ArrayDeque<CSVRecord> records = new ArrayDeque<>();
    private StringBuilder field = new StringBuilder();

    private final char[] lineDelimiter = new char[]{CSVQuoteStrategy.CARRIAGE_RETURN.getChar(), CSVQuoteStrategy.LINE_FEED.getChar()};

    private final char[] fieldSeparator = new char[]{','};
    private final char quote = '"';

    public CSVReader(InternalBuffer buffer) {
        this.buffer = buffer;
    }

    public ArrayDeque<CSVRecord> getRecords() {
        return records;
    }


    /**
     * 1. read -- 가칭 buffer 읽기 일단은 한글로
     *      역할 - buffer에서 record 단위로 char[] 생성 이후 new Record 생성 이후 콜백 Thread로
     */


    public void 레코드분리(){
        while (buffer.readerHasRemaining()){
            buffer.fillBuffer();
            레코드파싱();
        }
    }

    public void 레코드파싱(){
        int lineDelimiterIndex = 0;
        int fieldSeparatorIndex = 0;
        boolean inQuote = false;
        char prev = 0 ,cur = 0;

        while (buffer.hasRemaining()){
            cur = buffer.get();

            if (!inQuote){
                if (lineDelimiterIndex == lineDelimiter.length - 1){
                    // buffer의 offset 부터
                    // position 까지
                    break;
                }
                if (cur == lineDelimiter[lineDelimiterIndex]) {
                    lineDelimiterIndex++;
                    continue;
                }

                if (cur == fieldSeparator[fieldSeparatorIndex]){
                    fieldSeparatorIndex++;
                    continue;
                }

                if (cur == quote){
                    // 인용구 시작
                    inQuote = !inQuote;
                    lineDelimiterIndex = 0;
                    fieldSeparatorIndex = 0;
                    continue;
                }
            }

            if (inQuote){
                if (prev == quote && cur == quote){
                    prev = 0;
                    continue;
                }

                if (prev == quote && cur == fieldSeparator[fieldSeparatorIndex]){
                    fieldSeparatorIndex++;
                    continue;
                }

                if (prev == quote && fieldSeparatorIndex == fieldSeparator.length - 1){
                    inQuote = !inQuote;
                    prev = 0;
                    continue;
                }
                prev = cur;
            }
        }
    }

    // TODO 아래는 기존
    public void read(){
        if (records.isEmpty()) records.add(new CSVRecord());
        while (buffer.readerHasRemaining()){
            buffer.fillBufferIfEmpty(null);
            if (!buffer.readerHasRemaining()) break;

            if (buffer.getOffsetChar() == '"'){
                quoteFieldParse();
            } else {
                normalFieldParse();
            }
        }

        if (field.length() > 0) records.getLast().getRecord().add(field.toString());
    }

    private void quoteFieldParse(){
        boolean inQuote = true;
        buffer.get(); // skip started double quote
        buffer.updateOffset();
        char prev = 0;
        char cur = 0;
        buffer.fillBufferIfEmpty(null);

        while (buffer.hasRemaining()){
            cur = buffer.get();
            if ((cur == ',' || cur == '\n') && !inQuote) break;
            if (prev == '"' && cur == '"'){
                buffer.updateOffset();
                inQuote = !inQuote;
                prev = cur = 0;
            } else if (cur == '"' && inQuote) {
                buffer.sliceToAppend(b -> fieldAppend(b, buffer.getLength()));
                buffer.updateOffset();
                inQuote = !inQuote;
                prev = cur;
            }

            buffer.fillBufferIfEmpty(b -> fieldAppend(b, buffer.getLength()));
        }

        field.delete(field.length() - 1, field.length());
        records.getLast().addField(field.toString());
        clearField();
        buffer.updateOffset();
        if (cur == '\n') records.addLast(new CSVRecord());
    }

    private void normalFieldParse(){
        while (buffer.hasRemaining()){
            char c = buffer.get();
            if (c == ',' || c == '\n'){
                buffer.sliceToAppend(b -> fieldAppend(b, buffer.getLength() - 1));
                records.getLast().addField(field.toString());
                clearField();
                buffer.updateOffset();
                if (c == '\n') records.add(new CSVRecord());
                return;
            }
            buffer.fillBufferIfEmpty(b -> fieldAppend(b, buffer.getLength()));
        }

        buffer.sliceToAppend(b -> fieldAppend(b, buffer.getLength()));
        buffer.updateOffset();
    }

    private void clearField(){
        field.setLength(0);
    }

    private void fieldAppend(CharBuffer buffer, int length){
        field.append(buffer.array(), this.buffer.getOffset(), length);
    }



    // TODO Internal Buffer
    private static class InternalBuffer {
        private final Reader reader;
        private final CharBuffer buffer;
        private State state = State.REMAINING;
        private int offset = 0; // 이전 pos

        public InternalBuffer(Reader reader, int bufferSize) {
            this.reader = reader;
            this.buffer = CharBuffer.allocate(bufferSize);
            this.buffer.flip();
        }

        public boolean readerHasRemaining(){
            return state == State.REMAINING;
        }

        public void updateOffset(){
            this.offset = getPosition();
        }
        public int getOffset() {
            return this.offset;
        }
        public int getPosition(){
            return this.buffer.position();
        }

        public int getLength(){
            return getPosition() - this.offset;
        }

        public boolean hasRemaining(){
            return this.buffer.hasRemaining();
        }
        public void fillBufferIfEmpty(Consumer<CharBuffer> action){
            if (this.buffer.hasRemaining()){
                return;
            }
            if (action != null) action.accept(buffer);
            fillBuffer();
        }
        public char get(){
            return this.buffer.get();
        }

        public char getOffsetChar(){
            return this.buffer.get(offset);
        }
        public void sliceToAppend(Consumer<CharBuffer> consumer){
            consumer.accept(this.buffer);
        }


        public void fillBuffer(){
            try{
                this.buffer.clear();
                this.state = reader.read(buffer) < 0 ? State.EOF : state;
                this.buffer.flip();
                this.updateOffset();
            }catch (IOException e){
                throw new UncheckedIOException(e);
            }
        }
    }

    enum State {
        REMAINING,
        EOF
    }

    public static class ReaderBuilder{

        private char quote;
        private char fieldSeparator;


        public CSVReader build(Reader reader){
            return new CSVReader(new InternalBuffer(reader, 8192));
        }
    }
    // TODO Builder

}
