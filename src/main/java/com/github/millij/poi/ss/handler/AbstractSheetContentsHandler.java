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

    private int currentRow = 0;
    private Map<String, Object> currentRowObj;


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
        this.currentRowObj = new HashMap<String, Object>();
    }

    @Override
    public void endRow(int rowNum) {
        // Callback
        this.afterRowEnd(rowNum, new HashMap<String, Object>(currentRowObj));
    }

    @Override
    public void cell(String cellRef, String cellVal, XSSFComment comment) {
        // Sanity Checks
        if (StringUtils.isEmpty(cellRef)) {
            LOGGER.error("Row[#] {} : Cell reference is empty - {}", currentRow, cellRef);
            return;
        }

        if (StringUtils.isEmpty(cellVal)) {
            LOGGER.warn("Row[#] {} - Cell[ref] formatted value is empty : {} - {}", currentRow, cellRef, cellVal);
            return;
        }

        // CellColRef
        String cellColRef = Spreadsheet.getCellColumnReference(cellRef);

        // Set the CellValue in the Map
        LOGGER.debug("cell - Saving Column value : {} - {}", cellColRef, cellVal);
        currentRowObj.put(cellColRef, cellVal);
    }

    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {
        // TODO Auto-generated method stub

    }

}
