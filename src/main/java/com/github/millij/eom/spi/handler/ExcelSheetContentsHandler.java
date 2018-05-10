package com.github.millij.eom.spi.handler;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.millij.eom.ExcelUtil;
import com.github.millij.eom.spi.IExcelBean;


public class ExcelSheetContentsHandler<T extends IExcelBean> implements SheetContentsHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelSheetContentsHandler.class);

    private final Class<T> entityType;
    private final int noOfRowsToSkip;

    private final int headerRow;
    private final boolean verifyHeader;

    private final Map<String, String> entityPropertyMapping;

    private final Map<String, String> headerCellMap;
    private final List<T> rowObjList;

    private int currentRow = 0;
    private T currentRowEntity;


    // Constructors
    // ------------------------------------------------------------------------

    public ExcelSheetContentsHandler(Class<T> entityType) {
        this(entityType, 0, 0, true);
    }

    public ExcelSheetContentsHandler(Class<T> entityType, int noOfRowsToSkip) {
        this(entityType, noOfRowsToSkip, 0, true);
    }

    public ExcelSheetContentsHandler(Class<T> entityType, int noOfRowsToSkip, int headerRow, boolean verifyHeader) {
        super();

        // Sanity checks
        if (entityType == null) {
            throw new IllegalArgumentException("ExcelSheetContentsHandler :: Entity Class type is Null");
        }

        this.entityType = entityType;
        this.noOfRowsToSkip = noOfRowsToSkip;

        this.headerRow = headerRow;
        this.verifyHeader = verifyHeader;

        this.entityPropertyMapping = ExcelUtil.getColumnToPropertyMap(entityType);
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



    // SheetContentsHandler Implementations
    // ------------------------------------------------------------------------

    @Override
    public void startRow(int rowNum) {
        this.currentRow = rowNum;
        if (rowNum == headerRow || rowNum < noOfRowsToSkip) {
            return;
        }

        // Create a new instance of entity for each row
        try {
            currentRowEntity = entityType.newInstance();
        } catch (Exception ex) {
            String errMsg = String.format("Error occured while creating new instance of - %s", entityType);
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

        // Add the current row entity to the Objects list
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

            this.saveRowCellValue(cellRef, cellVal);
        }
    }

    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {
        // TODO Auto-generated method stub

    }


    // Private Methods
    // ------------------------------------------------------------------------

    private void verifySheetHeader() {
        // TODO Auto-generated method stub
    }

    private void saveHeaderCellValue(String cellRef, String cellValue) {
        String cellColRef = ExcelUtil.getCellColumnReference(cellRef);
        headerCellMap.put(cellColRef, cellValue);
    }

    private void saveRowCellValue(String cellRef, String cellValue) {
        // Now set the value in the row entity
        String cellColRef = ExcelUtil.getCellColumnReference(cellRef);
        String cellColName = headerCellMap.get(cellColRef);
        if (cellColName == null) {
            return;
        }

        String propName = this.entityPropertyMapping.get(cellColName);
        if (StringUtils.isEmpty(propName)) {
            LOGGER.debug("Row[#] {} : No mathching property found for column[name] - {}", currentRow, cellColName);
            return;
        }

        // Set the property value in the current row object bean
        try {
            BeanUtils.setProperty(currentRowEntity, propName, cellValue);
        } catch (Exception ex) {
            String errMsg = String.format("Error Setting bean property[name] - %s, value - %s : %s", propName, cellValue, ex.getMessage());
            LOGGER.error(errMsg, ex);
        }
    }


}
