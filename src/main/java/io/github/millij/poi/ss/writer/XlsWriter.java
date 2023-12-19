package io.github.millij.poi.ss.writer;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;


/**
 * {@link SpreadsheetWriter} implementation for XLS files.
 * 
 * @since 3.0
 */
public class XlsWriter extends AbstractSpreadsheetWriter {


    // Constructors

    public XlsWriter() {
        super(new HSSFWorkbook());
    }


    // Methods
    // ------------------------------------------------------------------------


}
