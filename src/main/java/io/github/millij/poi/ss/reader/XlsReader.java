package io.github.millij.poi.ss.reader;

import static io.github.millij.poi.util.Beans.isInstantiableType;

import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Objects;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import io.github.millij.poi.SpreadsheetReadException;
import io.github.millij.poi.ss.handler.RowListener;
import io.github.millij.poi.ss.model.Column;
import io.github.millij.poi.util.Spreadsheet;


/**
 * Reader implementation of {@link Workbook} for an POIFS file (.xls).
 * 
 * @see XlsxReader
 */
public class XlsReader extends AbstractSpreadsheetReader {

    private static final Logger LOGGER = LoggerFactory.getLogger(XlsReader.class);


    // Constructor

    public XlsReader() {
        this(0);
    }

    public XlsReader(final int headerRowIdx) {
        this(headerRowIdx, Integer.MAX_VALUE);
    }

    public XlsReader(final int headerRowIdx, final int lastRowIdx) {
        super(headerRowIdx, lastRowIdx);
    }


    // WorkbookReader Impl
    // ------------------------------------------------------------------------

    @Override
    public <T> void read(final Class<T> beanClz, final InputStream is, final RowListener<T> listener)
            throws SpreadsheetReadException {
        // Sanity checks
        if (!isInstantiableType(beanClz)) {
            throw new IllegalArgumentException("XlsReader :: Invalid bean type passed !");
        }

        try {
            final HSSFWorkbook wb = new HSSFWorkbook(is);
            final int sheetCount = wb.getNumberOfSheets();
            LOGGER.debug("Total no. of sheets found in HSSFWorkbook : #{}", sheetCount);

            // Iterate over sheets
            for (int i = 0; i < sheetCount; i++) {
                final HSSFSheet sheet = wb.getSheetAt(i);
                LOGGER.debug("Processing HSSFSheet at No. : {}", i);

                // Process Sheet
                this.processSheet(beanClz, sheet, listener);
            }

            // Close workbook
            wb.close();
        } catch (Exception ex) {
            String errMsg = String.format("Error reading HSSFSheet, to %s : %s", beanClz, ex.getMessage());
            LOGGER.error(errMsg, ex);
            throw new SpreadsheetReadException(errMsg, ex);
        }
    }

    @Override
    public <T> void read(final Class<T> beanClz, final InputStream is, final int sheetNo, final RowListener<T> listener)
            throws SpreadsheetReadException {
        // Sanity checks
        if (!isInstantiableType(beanClz)) {
            throw new IllegalArgumentException("XlsReader :: Invalid bean type passed !");
        }

        try {
            final HSSFWorkbook wb = new HSSFWorkbook(is);
            final HSSFSheet sheet = wb.getSheetAt(sheetNo - 1); // subtract 1 as Workbook follows 0-based index

            // Process Sheet
            this.processSheet(beanClz, sheet, listener);

            // Close workbook
            wb.close();
        } catch (Exception ex) {
            String errMsg = String.format("Error reading sheet %d, to %s : %s", sheetNo, beanClz, ex.getMessage());
            LOGGER.error(errMsg, ex);
            throw new SpreadsheetReadException(errMsg, ex);
        }
    }


    //
    // Read to Map

    @Override
    public void read(final InputStream is, final RowListener<Map<String, Object>> listener)
            throws SpreadsheetReadException {
        try {
            final HSSFWorkbook wb = new HSSFWorkbook(is);
            final int sheetCount = wb.getNumberOfSheets();
            LOGGER.debug("Total no. of sheets found in HSSFWorkbook : #{}", sheetCount);

            // Iterate over sheets
            for (int i = 0; i < sheetCount; i++) {
                final HSSFSheet sheet = wb.getSheetAt(i);
                LOGGER.debug("Processing HSSFSheet at No. : {}", i);

                // Process Sheet
                this.processSheet(sheet, listener);
            }

            // Close workbook
            wb.close();
        } catch (Exception ex) {
            String errMsg = String.format("Error reading HSSFSheet, to Map : %s", ex.getMessage());
            LOGGER.error(errMsg, ex);
            throw new SpreadsheetReadException(errMsg, ex);
        }
    }

    @Override
    public void read(final InputStream is, final int sheetNo, final RowListener<Map<String, Object>> listener)
            throws SpreadsheetReadException {
        try {
            final HSSFWorkbook wb = new HSSFWorkbook(is);
            final HSSFSheet sheet = wb.getSheetAt(sheetNo - 1); // subtract 1 as Workbook follows 0-based index

            // Process Sheet
            this.processSheet(sheet, listener);

            // Close workbook
            wb.close();
        } catch (Exception ex) {
            String errMsg = String.format("Error reading sheet %d, to Map : %s", sheetNo, ex.getMessage());
            LOGGER.error(errMsg, ex);
            throw new SpreadsheetReadException(errMsg, ex);
        }
    }


    //
    // Protected Methods
    // ------------------------------------------------------------------------

    protected void processSheet(final HSSFSheet sheet, final RowListener<Map<String, Object>> eventHandler) {
        // Header column - name mapping
        final HSSFRow headerRowObj = sheet.getRow(headerRowIdx);
        final Map<String, String> headerCellRefsMap = this.asHeaderNameToCellRefMap(headerRowObj, false);

        final Iterator<Row> rows = sheet.rowIterator();
        while (rows.hasNext()) {
            // Process Row Data
            final HSSFRow row = (HSSFRow) rows.next();
            final int rowNum = row.getRowNum();

            // Skip rows before Header ROW and after Last ROW
            if (rowNum < headerRowIdx || rowNum > lastRowIdx) {
                continue;
            }

            final Map<String, Object> rowDataMap = this.extractRowDataAsMap(row);
            if (rowDataMap == null || rowDataMap.isEmpty()) {
                continue;
            }

            // Row data as Bean
            final Map<String, Object> rowBean = Spreadsheet.rowAsMap(headerCellRefsMap, rowDataMap);
            eventHandler.row(rowNum, rowBean);
        }
    }

    protected <T> void processSheet(final Class<T> beanClz, final HSSFSheet sheet, final RowListener<T> eventHandler) {
        // Header column - name mapping
        final HSSFRow headerRowObj = sheet.getRow(headerRowIdx);
        final Map<String, String> headerCellRefsMap = this.asHeaderNameToCellRefMap(headerRowObj, true);

        // Bean Properties - column name mapping
        final Map<String, Column> propColumnMap = Spreadsheet.getPropertyToColumnDefMap(beanClz);

        final Iterator<Row> rows = sheet.rowIterator();
        while (rows.hasNext()) {
            // Process Row Data
            final HSSFRow row = (HSSFRow) rows.next();
            final int rowNum = row.getRowNum();

            // Skip rows before Header ROW and after Last ROW
            if (rowNum < headerRowIdx || rowNum > lastRowIdx) {
                continue;
            }

            final Map<String, Object> rowDataMap = this.extractRowDataAsMap(row);
            if (rowDataMap == null || rowDataMap.isEmpty()) {
                continue;
            }

            // Row data as Bean
            final T rowBean = Spreadsheet.rowAsBean(beanClz, propColumnMap, headerCellRefsMap, rowDataMap);
            eventHandler.row(rowNum, rowBean);
        }
    }


    // Private Methods
    // ------------------------------------------------------------------------

    private Map<String, String> asHeaderNameToCellRefMap(final HSSFRow headerRow, final boolean normalizeHeaderName) {
        // Sanity checks
        if (Objects.isNull(headerRow)) {
            return new HashMap<>();
        }

        final Map<String, String> headerCellRefs = new HashMap<>();

        final Iterator<Cell> cells = headerRow.cellIterator();
        while (cells.hasNext()) {
            final HSSFCell cell = (HSSFCell) cells.next();
            final int cellColIdx = cell.getColumnIndex();
            final String cellColRef = String.valueOf(cellColIdx);

            // Cell Value
            final Object header = this.getCellValue(cell);

            final String rawHeaderName = Objects.isNull(header) ? "" : String.valueOf(header);
            final String headerName = normalizeHeaderName ? Spreadsheet.normalize(rawHeaderName) : rawHeaderName;
            headerCellRefs.put(headerName, cellColRef);
        }

        return headerCellRefs;
    }

    private Map<String, Object> extractRowDataAsMap(final HSSFRow row) {
        // Sanity checks
        if (row == null) {
            return new HashMap<>();
        }

        final Map<String, Object> rowDataMap = new HashMap<>();

        final Iterator<Cell> cells = row.cellIterator();
        while (cells.hasNext()) {
            final HSSFCell cell = (HSSFCell) cells.next();
            final int cellColIdx = cell.getColumnIndex();
            final String cellColRef = String.valueOf(cellColIdx);

            // Cell Value
            final Object cellVal = this.getCellValue(cell);

            rowDataMap.put(cellColRef, cellVal);
        }

        return rowDataMap;
    }

    private Object getCellValue(final HSSFCell cell) {
        // Sanity checks
        if (Objects.isNull(cell)) {
            return null;
        }

        // Fetch value by CellType
        final Object cellVal;
        switch (cell.getCellType()) {
            case STRING:
                cellVal = cell.getStringCellValue();
                break;
            case NUMERIC:
                cellVal = cell.getNumericCellValue();
                break;
            case BOOLEAN:
                cellVal = cell.getBooleanCellValue();
                break;
            case FORMULA:
            case BLANK:
            case ERROR:
            default:
                cellVal = null;
                break;
        }

        return cellVal;
    }


}
