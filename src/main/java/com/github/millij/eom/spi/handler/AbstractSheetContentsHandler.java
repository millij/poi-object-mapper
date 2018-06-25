package com.github.millij.eom.spi.handler;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.millij.eom.ExcelUtil;


public abstract class AbstractSheetContentsHandler<T> implements SheetContentsHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(AbstractSheetContentsHandler.class);

    private final int noOfRowsToSkip;

    private final int headerRow;
    private final boolean verifyHeader;

    private final Map<String, String> headerCellMap;
    private final List<T> rowObjList;

    private int currentRow = 0;
    private T currentRowEntity;


    // Constructors
    // ------------------------------------------------------------------------

    public AbstractSheetContentsHandler(int noOfRowsToSkip) {
        this(noOfRowsToSkip, 0, true);
    }

    public AbstractSheetContentsHandler(int noOfRowsToSkip, int headerRow, boolean verifyHeader) {
        super();

        this.noOfRowsToSkip = noOfRowsToSkip;

        this.headerRow = headerRow;
        this.verifyHeader = verifyHeader;

        this.headerCellMap = new HashMap<String, String>();
        this.rowObjList = new ArrayList<T>();
    }


    // Methods
    // ------------------------------------------------------------------------

    /**
     * Returns the List of Object that read from the Excel Sheet.
     * 
     * @return Objects List
     */
    public List<T> getRowsAsObjects() {
        return new ArrayList<T>(rowObjList);
    }



    // Abstract

    public abstract T newEntityInstance();

    public abstract void saveRowCellValue(T currentRowObj, String cellColName, String cellValue);


    // SheetContentsHandler Implementations
    // ------------------------------------------------------------------------

    @Override
    public void startRow(int rowNum) {
        this.currentRow = rowNum;
        if (rowNum == headerRow || rowNum < noOfRowsToSkip) {
            return;
        }

        // Create a new instance of ExcelBean for each row
        try {
            currentRowEntity = newEntityInstance();
        } catch (Exception ex) {
            String errMsg = String.format("Error occured while creating new instance entity instance");
            LOGGER.error(errMsg, ex);
        }
    }

    @Override
    public void endRow(int rowNum) {
        if (this.headerRow == this.currentRow && verifyHeader) {
            this.verifySheetHeader();
        }

        if (currentRow < noOfRowsToSkip) {
            return;
        }

        // Add the current row Object to the Objects list
        if (currentRowEntity != null) {
            this.rowObjList.add(currentRowEntity);
        }

        // reset current Row object to null
        currentRowEntity = null;
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
            if (currentRowEntity == null) {
                return;
            }

            // ColumnName
            String cellColRef = ExcelUtil.getCellColumnReference(cellRef);
            String cellColName = headerCellMap.get(cellColRef);

            this.saveRowCellValue(currentRowEntity, cellColName, cellVal);
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
        String cellColRef = ExcelUtil.getCellColumnReference(cellRef);
        headerCellMap.put(cellColRef, cellValue);
    }


}
