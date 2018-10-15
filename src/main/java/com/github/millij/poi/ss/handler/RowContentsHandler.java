package com.github.millij.poi.ss.handler;

import java.lang.reflect.InvocationTargetException;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.millij.poi.util.Spreadsheet;


public class RowContentsHandler<T> extends AbstractSheetContentsHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(RowContentsHandler.class);

    private final Class<T> beanClz;
    private final Map<String, String> beanPropertyMap;

    private final int headerRow;
    private final Map<String, String> headerCellMap;

    private final RowListener<T> rowListener;


    // Constructors
    // ------------------------------------------------------------------------

    public RowContentsHandler(Class<T> beanClz, RowListener<T> rowListener) {
        this(beanClz, rowListener, 0);
    }

    public RowContentsHandler(Class<T> beanClz, RowListener<T> rowListener, int headerRow) {
        super();

        this.beanClz = beanClz;
        this.beanPropertyMap = Spreadsheet.getColumnToPropertyMap(beanClz);

        this.headerRow = headerRow;
        this.headerCellMap = new HashMap<String, String>();

        this.rowListener = rowListener;
    }


    // AbstractSheetContentsHandler Methods
    // ------------------------------------------------------------------------

    @Override
    void beforeRowStart(int rowNum) {
        // Row Callback
        // rowListener.beforeRow(rowNum);
    }


    @Override
    void afterRowEnd(int rowNum, Map<String, Object> rowDataMap) {
        // Sanity Checks
        if (rowDataMap == null || rowDataMap.isEmpty()) {
            return;
        }

        // Skip rows before header row
        if (rowNum < headerRow) {
            return;
        }

        if (rowNum == headerRow) {
            final Map<String, String> headerMap = this.prepareHeaderMap(rowNum, rowDataMap);
            headerCellMap.putAll(headerMap);
            return;
        }

        try {
            // Create new Instance
            T rowObj = beanClz.newInstance();

            // Fill in the datat
            for (String colRef : rowDataMap.keySet()) {
                String cellColName = headerCellMap.get(colRef);
                Object cellValue = rowDataMap.get(colRef);

                String propName = this.beanPropertyMap.get(cellColName);
                if (StringUtils.isEmpty(propName)) {
                    LOGGER.debug("Row[#] {} : No mathching property found for column[name] - {}", rowNum, colRef);
                    return;
                }

                // Set the property value in the current row object bean
                try {
                    BeanUtils.setProperty(rowObj, propName, cellValue);
                } catch (IllegalAccessException | InvocationTargetException ex) {
                    String errMsg = String.format("Failed to set bean property - %s, value - %s", propName, cellValue);
                    LOGGER.error(errMsg, ex);
                }
            }

            // Row Callback
            rowListener.row(rowNum, rowObj);
        } catch (Exception ex) {
            String errMsg = String.format("Error while creating bean - %s, from - %s", beanClz, rowDataMap);
            LOGGER.error(errMsg, ex);
        }
    }


    // Private Methods
    // ------------------------------------------------------------------------

    private Map<String, String> prepareHeaderMap(int rowNo, Map<String, Object> rowDataMap) {
        // Sanity checks
        if (rowDataMap == null || rowDataMap.isEmpty()) {
            String errMsg = String.format("Invalid Header data found - Row #%d", rowNo);
            throw new RuntimeException(errMsg);
        }

        final Map<String, String> headerMap = new HashMap<String, String>();
        for (String collRef : rowDataMap.keySet()) {
            Object colName = rowDataMap.get(collRef);
            if (colName != null) {
                headerMap.put(collRef, String.valueOf(colName));
            }
        }

        LOGGER.debug("Header DataMap prepared : {}", headerMap);
        return headerMap;
    }


}
