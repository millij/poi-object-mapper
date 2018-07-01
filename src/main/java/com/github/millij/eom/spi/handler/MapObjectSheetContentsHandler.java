package com.github.millij.eom.spi.handler;

import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;



public class MapObjectSheetContentsHandler extends AbstractSheetContentsHandler<Map<String, Object>> {

    private static final Logger LOGGER = LoggerFactory.getLogger(MapObjectSheetContentsHandler.class);


    // Constructors
    // ------------------------------------------------------------------------

    public MapObjectSheetContentsHandler() {
        super(0, 0, true);
    }

    public MapObjectSheetContentsHandler(int noOfRowsToSkip, int headerRow, boolean verifyHeader) {
        super(noOfRowsToSkip, headerRow, verifyHeader);
    }


    // Abstract Method Implementations
    // ------------------------------------------------------------------------

    @Override
    public Map<String, Object> newEntityInstance() {
        return new HashMap<String, Object>();
    }

    @Override
    public void saveRowCellValue(Map<String, Object> currentRowObj, String cellColName, String cellValue) {
        // Now set the value in the row Object
        if (StringUtils.isEmpty(cellColName)) {
            return;
        }

        LOGGER.debug("saveRowCellValue - Cell Column & value : {} - {}", cellColName, cellValue);

        // Set the CellValue in the Map
        currentRowObj.put(cellColName, cellValue);
    }



}
