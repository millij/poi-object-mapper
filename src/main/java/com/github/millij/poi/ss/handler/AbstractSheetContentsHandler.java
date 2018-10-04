package com.github.millij.poi.ss.handler;

import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.millij.poi.util.Spreadsheet;


public abstract class AbstractSheetContentsHandler implements SheetContentsHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(AbstractSheetContentsHandler.class);

    private final int noOfRowsToSkip;

    private final int headerRow;
    private final boolean verifyHeader;

    private final Map<String, String> headerCellMap;

    private int currentRow = 0;
    private Map<String, Object> currentRowObj;


    // Constructors
    // ------------------------------------------------------------------------

    AbstractSheetContentsHandler(int noOfRowsToSkip, int headerRow, boolean verifyHeader) {
        super();

        this.noOfRowsToSkip = noOfRowsToSkip;

        this.headerRow = headerRow;
        this.verifyHeader = verifyHeader;

        this.headerCellMap = new HashMap<String, String>();
    }


    // Methods
    // ------------------------------------------------------------------------

    // Abstract

    abstract void beforeRowStart(int rowNum);

    abstract void afterRowEnd(int rowNum, Map<String, Object> rowObj);



    // SheetContentsHandler Implementations
    // ------------------------------------------------------------------------

    @Override
    public void startRow(int rowNum) {
        // Callback
        this.beforeRowStart(rowNum);

        // Handle row
        this.currentRow = rowNum;
        if (rowNum == headerRow || rowNum < noOfRowsToSkip) {
            return;
        }

        // Init Row Object
        this.currentRowObj = new HashMap<String, Object>();
    }

    @Override
    public void endRow(int rowNum) {
        if (this.currentRow == this.headerRow  && verifyHeader) {
            this.verifySheetHeader();
            return;
        }

        if (currentRow < noOfRowsToSkip) {
            return;
        }

        // Callback
        this.afterRowEnd(rowNum, new HashMap<String, Object>(currentRowObj));
    }

    @Override
    public void cell(String cellRef, String cellVal, XSSFComment comment) {
        // Sanity Checks
        if (this.currentRow < noOfRowsToSkip) {
            return;
        }

        if (StringUtils.isEmpty(cellRef)) {
            LOGGER.error("Row[#] {} : Cell reference is empty - {}", currentRow, cellRef);
            return;
        }

        if (StringUtils.isEmpty(cellVal)) {
            LOGGER.warn("Row[#] {} - Cell[ref] formatted value is empty : {} - {}", currentRow, cellRef, cellVal);
            return;
        }

        // Handle the Header Row
        if (this.verifyHeader && this.headerRow == this.currentRow) {
            this.saveHeaderCellValue(cellRef, cellVal);
        } else {
            // ColumnName
            String cellColRef = Spreadsheet.getCellColumnReference(cellRef);
            String cellColName = headerCellMap.get(cellColRef);

            // Set the CellValue in the Map
            LOGGER.debug("cell - Saving Column value : {} - {}", cellColName, cellVal);
            currentRowObj.put(cellColName, cellVal);
        }
    }

    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {
        // TODO Auto-generated method stub

    }


    // Protected Methods
    // ------------------------------------------------------------------------

    protected void verifySheetHeader() {
        // TODO Auto-generated method stub
    }

    protected void saveHeaderCellValue(String cellRef, String cellValue) {
        String cellColRef = Spreadsheet.getCellColumnReference(cellRef);
        headerCellMap.put(cellColRef, cellValue);
    }



}
