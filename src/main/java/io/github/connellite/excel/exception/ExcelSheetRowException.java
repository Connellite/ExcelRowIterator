package io.github.connellite.excel.exception;

/**
 * Unchecked exception thrown when reading a row from a sheet iterator fails.
 */
public class ExcelSheetRowException extends RuntimeException {

    public ExcelSheetRowException(String message) {
        super(message);
    }

    public ExcelSheetRowException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelSheetRowException(Throwable cause) {
        super(cause);
    }
}
