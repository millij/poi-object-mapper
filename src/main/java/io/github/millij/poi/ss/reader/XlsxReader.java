package io.github.millij.poi.ss.reader;

import static io.github.millij.poi.util.Beans.isInstantiableType;



import io.github.millij.poi.SpreadsheetReadException;
import io.github.millij.poi.ss.handler.RowContentsHandler;
import io.github.millij.poi.ss.handler.RowListener;
import io.github.millij.poi.ss.model.annotations.SheetColumn;
import io.github.millij.poi.ss.writer.SpreadsheetWriter;
import io.github.millij.poi.util.Beans;
import io.github.millij.poi.util.Spreadsheet;

import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
// import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;


/**
 * Reader impletementation of {@link Workbook} for an OOXML .xlsx file. This implementation is suitable for low memory
 * sax parsing or similar.
 * 
 * @see XlsReader
 */
public class XlsxReader extends AbstractSpreadsheetReader {

    private static final Logger LOGGER = LoggerFactory.getLogger(XlsxReader.class);

    // Constructor

    public XlsxReader() {
        super();
    }


    // SpreadsheetReader Impl
    // ------------------------------------------------------------------------

    @Override
    public <T> void read(Class<T> beanClz, InputStream is, RowListener<T> listener) throws SpreadsheetReadException {
        // Sanity checks
        if (!isInstantiableType(beanClz)) {
            throw new IllegalArgumentException("XlsxReader_ftDate :: Invalid bean type passed !");
        }

        try {
            final XSSFWorkbook wb = new XSSFWorkbook(is);
            final int sheetCount = wb.getNumberOfSheets();
            LOGGER.debug("Total no. of sheets found in HSSFWorkbook : #{}", sheetCount);

            // Iterate over sheets
            for (int i = 0; i < sheetCount; i++) {
                final XSSFSheet sheet = wb.getSheetAt(i);
                LOGGER.debug("Processing HSSFSheet at No. : {}", i);

                // Process Sheet
                this.processSheet(beanClz, sheet, 0, listener);
            }

            // Close workbook
            wb.close();


        } catch (Exception ex) {
            String errMsg = String.format("Error reading sheet data, to Bean %s : %s", beanClz, ex.getMessage());
            LOGGER.error(errMsg, ex);
            throw new SpreadsheetReadException(errMsg, ex);
        }
    }

    @Override
    public <T> void read(Class<T> beanClz, InputStream is, int sheetNo, RowListener<T> listener)
            throws SpreadsheetReadException {
        // Sanity checks
        if (!isInstantiableType(beanClz)) {
            throw new IllegalArgumentException("XlsReader :: Invalid bean type passed !");
        }

        try {
            final XSSFWorkbook wb = new XSSFWorkbook(is);
            final XSSFSheet sheet = wb.getSheetAt(sheetNo);

            // Process Sheet
            this.processSheet(beanClz, sheet, 0, listener);

            // Close workbook
            wb.close();

        } catch (Exception ex) {
            String errMsg = String.format("Error reading sheet %d, to Bean %s : %s", sheetNo, beanClz, ex.getMessage());
            LOGGER.error(errMsg, ex);
            throw new SpreadsheetReadException(errMsg, ex);
        }
    }

    // Process Sheet

    protected <T> void processSheet(Class<T> beanClz, XSSFSheet sheet, int headerRowNo, RowListener<T> eventHandler) {
        // Header column - name mapping
        XSSFRow headerRow = sheet.getRow(headerRowNo);
        final Map<Integer, String> headerMap = this.extractCellHeaderMap(headerRow);

        // Bean Properties - column name mapping
        final Map<String, String> cellPropMapping = Spreadsheet.getColumnToPropertyMap(beanClz);

        Iterator<Row> rows = sheet.rowIterator();
        while (rows.hasNext()) {
            // Process Row Data
            XSSFRow row = (XSSFRow) rows.next();
            final int rowNum = row.getRowNum();
            if (rowNum == 0) {
                continue; // Skip Header row
            }

            final Map<String, Object> rowDataMap = this.extractRowDataAsMap(beanClz, row, headerMap);
            if (rowDataMap == null || rowDataMap.isEmpty()) {
                continue;
            }

            // Row data as Bean
            T rowBean = Spreadsheet.rowAsBean(beanClz, cellPropMapping, rowDataMap);
            eventHandler.row(rowNum, rowBean);
        }
    }


    // Private Methods

    private Map<Integer, String> extractCellHeaderMap(XSSFRow headerRow) {
        // Sanity checks
        if (headerRow == null) {
            return new HashMap<Integer, String>();
        }

        final Map<Integer, String> cellHeaderMap = new HashMap<Integer, String>();

        Iterator<Cell> cells = headerRow.cellIterator();
        while (cells.hasNext()) {
            XSSFCell cell = (XSSFCell) cells.next();

            final int cellCol = cell.getColumnIndex();

            // Process cell value
            switch (cell.getCellTypeEnum()) {
                case STRING:
                    cellHeaderMap.put(cellCol, cell.getStringCellValue());
                    break;
                case NUMERIC:

                    cellHeaderMap.put(cellCol, String.valueOf(cell.getNumericCellValue()));
                    break;

                case BOOLEAN:
                    cellHeaderMap.put(cellCol, String.valueOf(cell.getBooleanCellValue()));
                    break;
                case FORMULA:
                case BLANK:
                case ERROR:
                    break;
                default:
                    break;
            }
        }

        return cellHeaderMap;
    }


    private Map<String, Object> extractRowDataAsMap(Class<?> beanClz, XSSFRow row,
            Map<Integer, String> columnHeaderMap) {
        // Sanity checks
        if (row == null) {
            return new HashMap<String, Object>();
        }

        final Map<String, Object> rowDataMap = new HashMap<String, Object>();

        Iterator<Cell> cells = row.cellIterator();
        while (cells.hasNext()) {
            XSSFCell cell = (XSSFCell) cells.next();

            final int cellCol = cell.getColumnIndex();
            final String cellColName = columnHeaderMap.get(cellCol);

            // Process cell value
            switch (cell.getCellTypeEnum()) {
                case STRING:
                    rowDataMap.put(cellColName, cell.getStringCellValue());
                    break;
                case NUMERIC:
                    if (DateUtil.isCellDateFormatted(cell)) {

                        // Checking Date or LocalDate

                        String headerType = AbstractSpreadsheetReader.getReturnType(beanClz, cellColName);
                        if (headerType.equals(Date.class.getName())) {
                            final Date date = cell.getDateCellValue();
                            rowDataMap.put(cellColName, date);
                            break;
                        }
                        if (headerType.equals(LocalDate.class.getName())) {

                            final Date ldate = cell.getDateCellValue();

                            // Convert Date to LocalDate
                            Instant instant = ldate.toInstant();
                            ZoneId zoneId = ZoneId.systemDefault();
                            ZonedDateTime zonedDateTime = instant.atZone(zoneId);
                            LocalDate localDate = zonedDateTime.toLocalDate();

                            rowDataMap.put(cellColName, localDate);
                            break;
                        }
                    }
                case BOOLEAN:
                    rowDataMap.put(cellColName, cell.getBooleanCellValue());
                    break;
                case FORMULA:
                case BLANK:
                case ERROR:
                    break;
                default:
                    break;
            }
        }

        return rowDataMap;
    }

}
