package com.github.millij.poi.ss.reader;

import java.io.File;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.millij.poi.ExcelReadException;


/**
 * Reader impletementation of {@link Workbook} for an POIFS file (.xls). 
 * 
 * @see XlsxReader
 */
public class XlsReader extends WorkbookReader {

    private static final Logger LOGGER = LoggerFactory.getLogger(XlsReader.class);


    // Constructor

    public XlsReader() {
        super(WORKBOOK_XLS, false);
    }


    // WorkbookReader Impl
    // ------------------------------------------------------------------------

    @Override
    public <EB> List<EB> read(File file, int sheetNo, Class<EB> beanClz) throws ExcelReadException {
        // Sanity checks
        if (beanClz == null) {
            throw new IllegalArgumentException("XlsReader :: Bean class type should not be null");
        }

        try {
            // TODO complete
            
            return null;
        } catch (Exception ex) {
            String errMsg =
                    String.format("Error reading sheet %d, ExcelBean %s : %s", sheetNo, beanClz, ex.getMessage());
            LOGGER.error(errMsg, ex);
            throw new ExcelReadException(errMsg, ex);
        }

    }




}
