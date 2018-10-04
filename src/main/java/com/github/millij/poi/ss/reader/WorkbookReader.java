package com.github.millij.poi.ss.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.millij.poi.ExcelReadException;


/**
 * A abstract implementation of {@link Workbook} Reader.
 */
abstract class WorkbookReader {

    private static final Logger LOGGER = LoggerFactory.getLogger(WorkbookReader.class);

    // Labels

    public final static String WORKBOOK_XLS = "xls";
    public final static String WORKBOOK_XLSX = "xlsx";


    // Properties

    private final String _wbType;
    private final boolean _streaming;


    // Constructors
    // ------------------------------------------------------------------------

    WorkbookReader(String workbookType, boolean readAsSteam) {
        super();

        if (!isValidWorkbookType(workbookType)) {
            throw new IllegalArgumentException("WorkbookReader :: invalid workbook type");
        }

        // init
        this._wbType = workbookType;
        this._streaming = readAsSteam;
    }


    // Methods
    // ------------------------------------------------------------------------

    public <T> List<T> read(File file, Class<T> beanType) throws IOException, ExcelReadException {
        // Sanity checks
        validateTypeReference(beanType);

        // Get the workbook instance
        Workbook workbook = this.getWorkBook(_wbType, file, _streaming);
        int noOfSheets = workbook.getNumberOfSheets();
        LOGGER.debug("Total no of Sheets found : " + noOfSheets);

        // Iterate over all Sheets
        final List<T> beans = new ArrayList<T>();
        for (int i = 0; i < noOfSheets; i++) {
            beans.addAll(this.read(file, i, beanType));
        }

        return beans;
    }


    // Abstract Methods
    // ------------------------------------------------------------------------

    abstract <T> List<T> read(File file, int sheetNo, Class<T> beanType) throws IOException, ExcelReadException;



    // Protected Methods
    // ------------------------------------------------------------------------

    protected Workbook getWorkBook(String wbType, File file, boolean streamRead) throws IOException {
        // Workbook type
        FileInputStream fis = new FileInputStream(file);
        if (wbType.equals(WORKBOOK_XLS)) {
            return new HSSFWorkbook(fis);
        } else if (wbType.equals(WORKBOOK_XLSX)) {
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            if (streamRead) {
                return new SXSSFWorkbook(wb);
            }

            return wb;
        }

        IOUtils.closeQuietly(fis);
        return null;
    }


    // Private Methods
    // ------------------------------------------------------------------------



    // Static Utilities
    // ------------------------------------------------------------------------

    public static boolean isValidWorkbookType(String workbookType) {
        // Empty String
        if (workbookType == null || "".equals(workbookType.trim())) {
            return false;
        }

        // xls or xlsx
        if (WORKBOOK_XLS.equals(workbookType) || WORKBOOK_XLSX.equals(workbookType)) {
            return true;
        }

        return false;
    }

    public static void validateTypeReference(Class<?> beanType) {
        // Sanity checks
        if (beanType == null) {
            throw new IllegalArgumentException("Bean Type Reference should not be null.");
        }

        // Concrete class

        // TODO complete
    }


}
