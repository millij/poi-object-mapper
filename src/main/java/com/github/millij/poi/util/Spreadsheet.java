package com.github.millij.poi.util;

import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.millij.poi.ss.model.SheetColumn;


public class Spreadsheet {

    private static final Logger LOGGER = LoggerFactory.getLogger(Spreadsheet.class);

    private Spreadsheet() {
        // Utility Class
    }


    public final static String EXTN_XLS = "xls";
    public final static String EXTN_XLSX = "xlsx";

    
    // Utilities
    // ------------------------------------------------------------------------

    public static String getCellColumnReference(String cellRef) {
        // Sanity checks
        if (StringUtils.isEmpty(cellRef)) {
            return "";
        }

        // Splits the Cell name and returns the column reference
        String cellColRef = cellRef.split("[0-9]*$")[0];
        return cellColRef;
    }


    // Bean Data Mapping

    public static Map<String, String> asRowDataMap(Object beanObj, List<String> headers) throws Exception {
        // Sanity checks
        if (beanObj == null || CollectionUtils.isEmpty(headers)) {
            return new HashMap<>();
        }

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
            if (!headers.contains(header)) {
                continue;
            }

            rowDataMap.put(header, getFieldValueAsString(beanObj, fieldName));
        }

        // Methods
        for (Method m : beanType.getDeclaredMethods()) {
            if (!m.isAnnotationPresent(SheetColumn.class)) {
                continue;
            }

            String fieldName = getFieldName(m);

            SheetColumn ec = m.getAnnotation(SheetColumn.class);
            String header = StringUtils.isEmpty(ec.value()) ? fieldName : ec.value();
            if (!headers.contains(header)) {
                continue;
            }

            rowDataMap.put(header, getFieldValueAsString(beanObj, fieldName));
        }

        return rowDataMap;
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
            String fieldName = f.getName();
            mapping.put(fieldName, fieldName);

            SheetColumn ec = f.getAnnotation(SheetColumn.class);
            if (ec != null && StringUtils.isNotEmpty(ec.value())) {
                mapping.put(fieldName, ec.value());
            }
        }

        // Methods
        Method[] methods = beanType.getDeclaredMethods();
        for (Method m : methods) {
            String fieldName = getFieldName(m);
            if (!mapping.containsKey(fieldName)) {
                mapping.put(fieldName, fieldName);
            }

            SheetColumn ec = m.getAnnotation(SheetColumn.class);
            if (ec != null && StringUtils.isNotEmpty(ec.value())) {
                mapping.put(fieldName, ec.value());
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

        final ArrayList<String> columnNames = new ArrayList<>(propToColumnMap.values());
        return columnNames;
    }


    // Private Utils
    // ------------------------------------------------------------------------

    private static String getFieldName(Method method) {
        // Sanity checks
        if (method == null) {
            return null;
        }

        String methodName = method.getName();
        return Introspector.decapitalize(methodName.substring(methodName.startsWith("is") ? 2 : 3));
    }

    private static String getFieldValueAsString(Object beanObj, String fieldName) throws Exception {
        // Sanity checks
        PropertyDescriptor pd = new PropertyDescriptor(fieldName, beanObj.getClass());
        Method getterMtd = pd.getReadMethod();

        Object value = getterMtd.invoke(beanObj);
        String cellValue = value != null ? value.toString() : "";

        return cellValue;
    }


}
