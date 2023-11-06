package org.celper.exception;

/**
 * The type Empty sheet exception.
 */
public class EmptySheetException extends ExcelException{
    /**
     * Instantiates a new Empty sheet exception.
     *
     * @param message the message
     */
    public EmptySheetException(String message) {
        super(message);
    }

    /**
     * Instantiates a new Empty sheet exception.
     *
     * @param message the message
     * @param cause   the cause
     */
    public EmptySheetException(String message, Throwable cause) {
        super(message, cause);
    }
}
