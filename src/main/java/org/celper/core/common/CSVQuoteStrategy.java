package org.celper.core.common;

public enum CSVQuoteStrategy {
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

    public boolean contains(String str) {
        return str.contains(charSequence);
    }
}
