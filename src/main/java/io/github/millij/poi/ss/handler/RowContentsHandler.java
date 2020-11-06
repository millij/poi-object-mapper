package io.github.millij.poi.ss.handler;

import io.github.millij.poi.util.Spreadsheet;

import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class RowContentsHandler<T> extends AbstractSheetContentsHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(RowContentsHandler.class);

    private final Class<T> beanClz;

    private final int headerRow;
    private final Map<String, String> cellPropertyMap;

    private final RowListener<T> rowListener;


    // Constructors
    // ------------------------------------------------------------------------

    public RowContentsHandler(Class<T> beanClz, RowListener<T> rowListener) {
        this(beanClz, rowListener, 0);
    }

    public RowContentsHandler(Class<T> beanClz, RowListener<T> rowListener, int headerRow) {
        super();

        this.beanClz = beanClz;

        this.headerRow = headerRow;
        this.cellPropertyMap = new HashMap<String, String>();

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
            cellPropertyMap.putAll(headerMap);
            return;
        }

        // Row As Bean
        T rowBean = Spreadsheet.rowAsBean(beanClz, cellPropertyMap, rowDataMap);

        // Row Callback
        try {
            rowListener.row(rowNum, rowBean);
        } catch (Exception ex) {
            String errMsg = String.format("Error calling listener callback  row - %d, bean - %s", rowNum, rowBean);
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

        final Map<String, String> colToBeanPropMap = Spreadsheet.getColumnToPropertyMap(beanClz);

        final Map<String, String> headerMap = new HashMap<String, String>();
        for (String collRef : rowDataMap.keySet()) {
            String colName = String.valueOf(rowDataMap.get(collRef)).trim();
            String propName = colToBeanPropMap.get(colName);
            if (StringUtils.isNotEmpty(propName)) {
                headerMap.put(collRef, String.valueOf(propName));
            }
        }

        LOGGER.debug("Header DataMap prepared : {}", headerMap);
        return headerMap;
    }


}
