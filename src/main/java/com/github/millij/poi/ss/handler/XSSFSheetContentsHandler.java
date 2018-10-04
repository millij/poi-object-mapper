package com.github.millij.poi.ss.handler;

import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.millij.poi.util.Spreadsheet;



public class XSSFSheetContentsHandler<T> extends AbstractSheetContentsHandler {

    private static final Logger LOGGER = LoggerFactory.getLogger(XSSFSheetContentsHandler.class);

    private final Class<T> beanType;
    private final List<T> rowObjects;

    private final Map<String, String> beanPropertyMapping;


    // Constructors
    // ------------------------------------------------------------------------

    public XSSFSheetContentsHandler(Class<T> beanType) {
        this(beanType, 0, 0, true);
    }

    public XSSFSheetContentsHandler(Class<T> beanType, int noOfRowsToSkip, int headerRow, boolean verifyHeader) {
        super(noOfRowsToSkip, headerRow, verifyHeader);

        this.beanType = beanType;
        this.rowObjects = new ArrayList<T>();

        this.beanPropertyMapping = Spreadsheet.getColumnToPropertyMap(beanType);
    }


    // Abstract Method Implementations
    // ------------------------------------------------------------------------

    @Override
    void beforeRowStart(int rowNum) {

    }

    @Override
    void afterRowEnd(int rowNum, Map<String, Object> rowDataMap) {
        // Sanity checks
        if (rowDataMap == null || rowDataMap.isEmpty()) {
            return;
        }

        try {
            // Create new Instance
            T rowObj = beanType.newInstance();

            // Fill in the datat
            for (String colName : rowDataMap.keySet()) {
                Object cellValue = rowDataMap.get(colName);

                String propName = this.beanPropertyMapping.get(colName);
                if (StringUtils.isEmpty(propName)) {
                    LOGGER.debug("Row[#] {} : No mathching property found for column[name] - {}", rowNum, colName);
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

            this.rowObjects.add(rowObj);

        } catch (Exception ex) {
            String errMsg = String.format("Error while creating bean - %s, from - %s", beanType, rowDataMap);
            LOGGER.error(errMsg, ex);
        }

    }


    // Getters and Setters
    // ------------------------------------------------------------------------

    public Class<T> getBeanType() {
        return beanType;
    }

    public List<T> getRowObjects() {
        return new ArrayList<T>(rowObjects);
    }


}
