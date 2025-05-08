package org.celper;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;


/**
 * 기본적으로 RFC 4180 문서에 정의를 따른다.
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.SOURCE)
public @interface CSVConfig {
    String lineDelimiter() default "";
    String fieldSeparator() default "";
    char quote() default '\0'; // 단일 문자열이 아닌 2개 이상의 문자열이 들어오면 예외 난다고 표시
    String comment() default "";

    enum DefaultCSVConfig{
        LINE_DELIMITER_CR("\n"), // linux
        LINE_DELIMITER_LF("\r"), // unix
        LINE_DELIMITER_CRLF("\r\n"), // windows
        FIELD_SEPARATOR(","), // field 나누는 방법
        QUOTE_STRATEGY("\""), // 인용구 및 Line Delimiter 를 감싸는 방법
        COMMENT_PREFIX("#"); // 주석의 시작에 대한 표기 방법

        private String str;

        DefaultCSVConfig(String str) {
            this.str = str;
        }

        public static String[] getDefaultCSVConfig(){
            return new String[]{
                    LINE_DELIMITER_CR.str,
                    FIELD_SEPARATOR.str,
                    QUOTE_STRATEGY.str,
                    COMMENT_PREFIX.str};
        }

        public String getStr() {
            return str;
        }
    }
}
