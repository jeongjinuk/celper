package org.celper.exception;

public class EmptySheetException extends ExcelException{
    public EmptySheetException(String message) {
        super(message);
    }

    public EmptySheetException(String message, Throwable cause) {
        super(message, cause);
    }
}
