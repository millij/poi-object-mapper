package io.github.millij.poi.ss.handler;

import java.util.HashMap;
import java.util.Map;
import java.util.Objects;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import io.github.millij.poi.ss.model.Column;
import io.github.millij.poi.util.Spreadsheet;
import io.github.millij.poi.util.Strings;


public class RowContentsHandler<T> extends AbstractSheetContentsHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(RowContentsHandler.class);

    private final Class<T> beanClz;
    private final Map<String, Column> beanPropColumnMap;

    private final RowListener<T> listener;

    private final int headerRowNum;
    private final Map<String, String> headerCellRefsMap;

    private final int lastRowNum;


    // Constructors
    // ------------------------------------------------------------------------

    public RowContentsHandler(Class<T> beanClz, RowListener<T> listener, int headerRowNum, int lastRowNum) {
        super();

        // init
        this.beanClz = beanClz;
        this.beanPropColumnMap = Spreadsheet.getPropertyToColumnDefMap(beanClz);

        this.listener = listener;

        this.headerRowNum = headerRowNum;
        this.headerCellRefsMap = new HashMap<>();

        this.lastRowNum = lastRowNum;
    }


    // AbstractSheetContentsHandler Methods
    // ------------------------------------------------------------------------

    @Override
    void beforeRowStart(final int rowNum) {
        try {
            // Row Callback
            listener.beforeRow(rowNum);
        } catch (Exception ex) {
            String errMsg = String.format("Error calling #beforeRow callback  row - %d", rowNum);
            LOGGER.error(errMsg, ex);
        }
    }


    @Override
    void afterRowEnd(final int rowNum, final Map<String, Object> rowDataMap) {
        // Sanity Checks
        if (Objects.isNull(rowDataMap) || rowDataMap.isEmpty()) {
            LOGGER.debug("INVALID Row data Passed - Row #{}", rowNum);
            return;
        }

        // Skip rows before Header ROW and after Last ROW
        if (rowNum < headerRowNum || rowNum > lastRowNum) {
            return;
        }

        // Process Header ROW
        if (rowNum == headerRowNum) {
            final Map<String, String> headerCellRefs = this.asHeaderNameToCellRefMap(rowDataMap);
            headerCellRefsMap.putAll(headerCellRefs);
            return;
        }

        // Check for Column Definitions before processing NON-Header ROWs

        // Row As Bean
        final T rowBean = Spreadsheet.rowAsBean(beanClz, beanPropColumnMap, headerCellRefsMap, rowDataMap);
        if (Objects.isNull(rowBean)) {
            LOGGER.debug("Unable to construct Row data Bean object - Row #{}", rowNum);
            return;
        }

        // Row Callback
        try {
            listener.row(rowNum, rowBean);
        } catch (Exception ex) {
            String errMsg = String.format("Error calling #row callback  row - %d, bean - %s", rowNum, rowBean);
            LOGGER.error(errMsg, ex);
        }
    }


    // Private Methods
    // ------------------------------------------------------------------------

    private Map<String, String> asHeaderNameToCellRefMap(final Map<String, Object> headerRowData) {
        // Sanity checks
        if (Objects.isNull(headerRowData) || headerRowData.isEmpty()) {
            return new HashMap<>();
        }

        // Get Bean Column definitions
        final Map<String, String> headerCellRefs = new HashMap<String, String>();
        for (final String colRef : headerRowData.keySet()) {
            final Object header = headerRowData.get(colRef);

            final String headerName = Objects.isNull(header) ? "" : String.valueOf(header);
            final String normalHeaderName = Strings.normalize(headerName);
            headerCellRefs.put(normalHeaderName, colRef);
        }

        LOGGER.debug("Header Name to Cell Refs : {}", headerCellRefs);
        return headerCellRefs;
    }


}
