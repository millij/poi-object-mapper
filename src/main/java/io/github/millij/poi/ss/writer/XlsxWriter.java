package io.github.millij.poi.ss.writer;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * {@link SpreadsheetWriter} implementation for XLSX files.
 * 
 * @since 3.0
 */
public class XlsxWriter extends AbstractSpreadsheetWriter {


    // Constructors

    public XlsxWriter() {
        super(new XSSFWorkbook());
    }


    // Methods
    // ------------------------------------------------------------------------


}
