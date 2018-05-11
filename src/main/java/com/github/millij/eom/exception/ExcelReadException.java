package com.github.millij.eom.exception;


public class ExcelReadException extends Exception {
    
    private static final long serialVersionUID = 1L;
    
    
    // Constructors
    // ------------------------------------------------------------------------

    public ExcelReadException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelReadException(String message) {
        super(message);
    }
    
}
