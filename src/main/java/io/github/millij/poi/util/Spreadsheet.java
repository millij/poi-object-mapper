package io.github.millij.poi.util;

import io.github.millij.poi.ss.model.annotations.SheetColumn;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Spreadsheet related utilites.
 */
public final class Spreadsheet {

    private static final Logger LOGGER = LoggerFactory.getLogger(Spreadsheet.class);

    private Spreadsheet() {
        // Utility Class
    }


    // Utilities
    // ------------------------------------------------------------------------

    /**
     * Splits the CellReference and returns only the column reference.
     * 
     * @param cellRef the cell reference value (ex. D3)
     * @return returns the column index "D" from the cell reference "D3"
     */
    public static String getCellColumnReference(String cellRef) {
        String cellColRef = cellRef.split("[0-9]*$")[0];
        return cellColRef;
    }



    // Bean :: Property Utils

    public static Map<String, String> getPropertyToColumnNameMap(Class<?> beanType) {
        // Sanity checks
        if (beanType == null) {
            throw new IllegalArgumentException("getColumnToPropertyMap :: Invalid ExcelBean type - " + beanType);
        }

        // Property to Column name Mapping
        final Map<String, String> mapping = new HashMap<String, String>();

        // Fields
        Field[] fields = beanType.getDeclaredFields();
        for (Field f : fields) {
            if (!f.isAnnotationPresent(SheetColumn.class)) {
                continue;
            }

            final String fieldName = f.getName();
            SheetColumn ec = f.getAnnotation(SheetColumn.class);
            if (ec != null) {
                mapping.put(fieldName, StringUtils.isBlank(ec.value()) ? fieldName : ec.value());
            }
        }

        // Methods
        Method[] methods = beanType.getDeclaredMethods();
        for (Method m : methods) {
            if (!m.isAnnotationPresent(SheetColumn.class)) {
                continue;
            }

            final String fieldName = Beans.getFieldName(m);
            SheetColumn ec = m.getAnnotation(SheetColumn.class);
            if (ec != null) {
                mapping.put(fieldName, StringUtils.isBlank(ec.value()) ? fieldName : ec.value());
            }
        }

        LOGGER.info("Bean property to Excel Column of - {} : {}", beanType, mapping);
        return Collections.unmodifiableMap(mapping);
    }

    public static Map<String, String> getColumnToPropertyMap(Class<?> beanType) {
        // Column to Property Mapping
        final Map<String, String> columnToPropMap = new HashMap<String, String>();

        // Bean Property to Column Mapping
        final Map<String, String> propToColumnMap = getPropertyToColumnNameMap(beanType);
        for (String prop : propToColumnMap.keySet()) {
            columnToPropMap.put(propToColumnMap.get(prop), prop);
        }

        LOGGER.info("Excel Column to property map of - {} : {}", beanType, columnToPropMap);
        return Collections.unmodifiableMap(columnToPropMap);
    }

    public static List<String> getColumnNames(Class<?> beanType) {
        // Bean Property to Column Mapping
        final Map<String, String> propToColumnMap = getPropertyToColumnNameMap(beanType);

        final Map<Integer, String> indexToPropMap = getIndexToPropertyMap(beanType);

        if (propToColumnMap.size() == indexToPropMap.size()) {

            final Set<Integer> indexSet = indexToPropMap.keySet();
            List<Integer> indexList = new ArrayList<Integer>(indexSet);
            Collections.sort(indexList);

            List<String> indexedColumns = new ArrayList<String>();
            for (Integer index : indexList) {
                indexedColumns.add(propToColumnMap.get(indexToPropMap.get(index)));
                if (index == -1) {
                    LOGGER.info("Writing One field : '{}' at first column as no index specified",
                            propToColumnMap.get(indexToPropMap.get(index)));
                }
            }
            return indexedColumns;
        }


        LOGGER.info("Failed to write headers in partially indexed order. Proceeded to write in random order");
        final ArrayList<String> columnNames = new ArrayList<>(propToColumnMap.values());
        return columnNames;
    }


    public static Map<Integer, String> getIndexToPropertyMap(Class<?> beanClz) {

        Map<Integer, String> indexToPropMap = new HashMap<Integer, String>();

        Field[] fields = beanClz.getDeclaredFields();
        for (Field f : fields) {
            if (!f.isAnnotationPresent(SheetColumn.class)) {
                continue;
            }
            final String fieldName = f.getName();
            SheetColumn ec = f.getDeclaredAnnotation(SheetColumn.class);
            indexToPropMap.put(ec.index(), fieldName);
        }

        Method[] methods = beanClz.getDeclaredMethods();
        for (Method m : methods) {
            if (!m.isAnnotationPresent(SheetColumn.class)) {
                continue;
            }
            final String fieldName = Beans.getFieldName(m);
            SheetColumn ec = m.getDeclaredAnnotation(SheetColumn.class);
            indexToPropMap.put(ec.index(), fieldName);
        }

        return indexToPropMap;
    }



    // Read from Bean : as Row Data
    // ------------------------------------------------------------------------

    public static Map<String, String> asRowDataMap(Object beanObj, List<String> colHeaders) throws Exception {
        // Excel Bean Type
        final Class<?> beanType = beanObj.getClass();

        // RowData map
        final Map<String, String> rowDataMap = new HashMap<String, String>();

        // Fields
        for (Field f : beanType.getDeclaredFields()) {
            if (!f.isAnnotationPresent(SheetColumn.class)) {
                continue;
            }

            String fieldName = f.getName();

            SheetColumn ec = f.getAnnotation(SheetColumn.class);
            String header = StringUtils.isEmpty(ec.value()) ? fieldName : ec.value();
            if (!colHeaders.contains(header)) {
                continue;
            }

            rowDataMap.put(header, Beans.getFieldValueAsString(beanObj, fieldName));
        }

        // Methods
        for (Method m : beanType.getDeclaredMethods()) {
            if (!m.isAnnotationPresent(SheetColumn.class)) {
                continue;
            }

            String fieldName = Beans.getFieldName(m);

            SheetColumn ec = m.getAnnotation(SheetColumn.class);
            String header = StringUtils.isEmpty(ec.value()) ? fieldName : ec.value();
            if (!colHeaders.contains(header)) {
                continue;
            }

            rowDataMap.put(header, Beans.getFieldValueAsString(beanObj, fieldName));
        }

        return rowDataMap;
    }



    // Write to Bean :: from Row data
    // ------------------------------------------------------------------------

    public static <T> T rowAsBean(Class<T> beanClz, Map<String, String> cellProperies, Map<String, Object> cellValues) {
        // Sanity checks
        if (cellValues == null || cellProperies == null) {
            return null;
        }

        try {
            // Create new Instance
            T rowBean = beanClz.newInstance();

            // Fill in the datat
            for (String cellName : cellProperies.keySet()) {
                String propName = cellProperies.get(cellName);
                if (StringUtils.isEmpty(propName)) {
                    LOGGER.debug("{} : No mathching property found for column[name] - {} ", beanClz, cellName);
                    continue;
                }

                Object propValue = cellValues.get(cellName);
                try {
                    // Set the property value in the current row object bean
                    BeanUtils.setProperty(rowBean, propName, propValue);
                } catch (IllegalAccessException | InvocationTargetException ex) {
                    String errMsg = String.format("Failed to set bean property - %s, value - %s", propName, propValue);
                    LOGGER.error(errMsg, ex);
                }
            }

            return rowBean;
        } catch (Exception ex) {
            String errMsg = String.format("Error while creating bean - %s, from - %s", beanClz, cellValues);
            LOGGER.error(errMsg, ex);
        }

        return null;
    }



}
