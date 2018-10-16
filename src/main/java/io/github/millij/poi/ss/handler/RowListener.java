package io.github.millij.poi.ss.handler;

import io.github.millij.poi.ss.reader.SpreadsheetReader;


/**
 * An abstract representation of Row level Callback for {@link SpreadsheetReader} implementations.
 */
public interface RowListener<T> {

    /**
     * This method will be called after every row by the {@link SpreadsheetReader} implementation.
     * 
     * @param rowNum the Row Number in the sheet. (indexed from 0)
     * @param rowObj
     */
    void row(int rowNum, T rowObj);

}
