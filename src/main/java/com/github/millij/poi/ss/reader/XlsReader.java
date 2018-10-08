package com.github.millij.poi.ss.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.millij.poi.ExcelReadException;
import com.github.millij.poi.util.Spreadsheet;


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

        final List<EB> sheetBeans = new ArrayList<EB>();
        try {
            InputStream ExcelFileToRead = new FileInputStream(file);
            HSSFWorkbook wb = new HSSFWorkbook(ExcelFileToRead);

            HSSFSheet sheet = wb.getSheetAt(sheetNo);
            HSSFRow headerRow = sheet.getRow(0);

            // Header column - name mapping
            Map<Integer, String> columnHeaderMap = this.extractColumnHeaderMap(headerRow);
            
            // Bean Properties - column name mapping
            Map<String, String> cellPropMapping = Spreadsheet.getColumnToPropertyMap(beanClz);

            Iterator<Row> rows = sheet.rowIterator();
            while (rows.hasNext()) {
                // Process Row Data
                HSSFRow row = (HSSFRow) rows.next();
                if (row.getRowNum() == 0) {
                    continue; // Skip Header row
                }

                Map<String, Object> rowDataMap = this.extractRowDataAsMap(row, columnHeaderMap);
                if (rowDataMap == null || rowDataMap.isEmpty()) {
                    continue;
                }

                // Row data as Bean
                EB rowBean = Spreadsheet.rowAsBean(beanClz, cellPropMapping, rowDataMap);
                if (rowBean != null) {
                    sheetBeans.add(rowBean);
                }
            }

            // Close workbook
            wb.close();

            return Collections.unmodifiableList(sheetBeans);

        } catch (Exception ex) {
            String errMsg = String.format("Error reading sheet %d, to %s : %s", sheetNo, beanClz, ex.getMessage());
            LOGGER.error(errMsg, ex);
            throw new ExcelReadException(errMsg, ex);
        }

    }



    // Private Methods
    // ------------------------------------------------------------------------

    private Map<Integer, String> extractColumnHeaderMap(HSSFRow headerRow) {
        // Sanity checks
        if (headerRow == null) {
            return new HashMap<Integer, String>();
        }

        final Map<Integer, String> cellHeaderMap = new HashMap<Integer, String>();

        Iterator<Cell> cells = headerRow.cellIterator();
        while (cells.hasNext()) {
            HSSFCell cell = (HSSFCell) cells.next();

            int cellCol = cell.getColumnIndex();

            // Process cell value
            switch (cell.getCellType()) {
                case HSSFCell.CELL_TYPE_STRING:
                    cellHeaderMap.put(cellCol, cell.getStringCellValue());
                    break;
                case HSSFCell.CELL_TYPE_NUMERIC:
                    cellHeaderMap.put(cellCol, String.valueOf(cell.getNumericCellValue()));
                    break;
                case HSSFCell.CELL_TYPE_BOOLEAN:
                    cellHeaderMap.put(cellCol, String.valueOf(cell.getBooleanCellValue()));
                    break;
                case HSSFCell.CELL_TYPE_FORMULA:
                case HSSFCell.CELL_TYPE_BLANK:
                case HSSFCell.CELL_TYPE_ERROR:
                    break;
                default:
                    break;
            }
        }

        return cellHeaderMap;
    }

    private Map<String, Object> extractRowDataAsMap(HSSFRow row, Map<Integer, String> columnHeaderMap) {
        // Sanity checks
        if (row == null) {
            return new HashMap<String, Object>();
        }

        final Map<String, Object> rowDataMap = new HashMap<String, Object>();

        Iterator<Cell> cells = row.cellIterator();
        while (cells.hasNext()) {
            HSSFCell cell = (HSSFCell) cells.next();

            int cellCol = cell.getColumnIndex();
            String cellColName = columnHeaderMap.get(cellCol);

            // Process cell value
            switch (cell.getCellType()) {
                case HSSFCell.CELL_TYPE_STRING:
                    rowDataMap.put(cellColName, cell.getStringCellValue());
                    break;
                case HSSFCell.CELL_TYPE_NUMERIC:
                    rowDataMap.put(cellColName, cell.getNumericCellValue());
                    break;
                case HSSFCell.CELL_TYPE_BOOLEAN:
                    rowDataMap.put(cellColName, cell.getBooleanCellValue());
                    break;
                case HSSFCell.CELL_TYPE_FORMULA:
                case HSSFCell.CELL_TYPE_BLANK:
                case HSSFCell.CELL_TYPE_ERROR:
                    break;
                default:
                    break;
            }
        }

        return rowDataMap;
    }



}
