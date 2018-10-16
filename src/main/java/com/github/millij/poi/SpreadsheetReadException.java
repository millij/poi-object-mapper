package com.github.millij.poi;


public class SpreadsheetReadException extends Exception {

    private static final long serialVersionUID = 1L;


    // Constructors
    // ------------------------------------------------------------------------

    public SpreadsheetReadException(String message, Throwable cause) {
        super(message, cause);
    }

    public SpreadsheetReadException(String message) {
        super(message);
    }

}
