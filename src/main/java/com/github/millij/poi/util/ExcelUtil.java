package com.github.millij.poi.util;

import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.File;
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

import com.github.millij.poi.spi.annotation.ExcelColumn;


public class ExcelUtil {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelUtil.class);

    private ExcelUtil() {
        // Utility Class
    }


    public final static String EXTN_XLS = "xls";
    public final static String EXTN_XLSX = "xlsx";

    
    // Utilities
    // ------------------------------------------------------------------------


    // File

    @Deprecated
    public static String getFileExtension(File inFile) {
        // Sanity checks
        if (inFile == null) {
            throw new IllegalArgumentException("getFileExtension :: file should not be null");
        }

        String filepath = inFile.getAbsolutePath();
        LOGGER.debug("Extracting extension of file : {}", filepath);

        int p = Math.max(filepath.lastIndexOf('/'), filepath.lastIndexOf('\\'));
        int i = filepath.lastIndexOf('.');
        if (i > p) {
            return filepath.substring(i + 1);
        }

        return null;
    }


    // Excel

    @Deprecated
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
        final Class<?> excelBeanType = beanObj.getClass();

        // RowData map
        final Map<String, String> rowDataMap = new HashMap<String, String>();

        // Fields
        for (Field f : excelBeanType.getDeclaredFields()) {
            if (!f.isAnnotationPresent(ExcelColumn.class)) {
                continue;
            }

            String fieldName = f.getName();

            ExcelColumn ec = f.getAnnotation(ExcelColumn.class);
            String header = StringUtils.isEmpty(ec.value()) ? fieldName : ec.value();
            if (!headers.contains(header)) {
                continue;
            }

            rowDataMap.put(header, getFieldValueAsString(beanObj, fieldName));
        }

        // Methods
        for (Method m : excelBeanType.getDeclaredMethods()) {
            if (!m.isAnnotationPresent(ExcelColumn.class)) {
                continue;
            }

            String fieldName = getFieldName(m);

            ExcelColumn ec = m.getAnnotation(ExcelColumn.class);
            String header = StringUtils.isEmpty(ec.value()) ? fieldName : ec.value();
            if (!headers.contains(header)) {
                continue;
            }

            rowDataMap.put(header, getFieldValueAsString(beanObj, fieldName));
        }

        return rowDataMap;
    }



    // Bean :: Property Utils

    public static Map<String, String> getPropertyToColumnNameMap(Class<?> excelBeanType) {
        // Sanity checks
        if (excelBeanType == null) {
            throw new IllegalArgumentException("getColumnToPropertyMap :: Invalid ExcelBean type - " + excelBeanType);
        }

        // Property to Column name Mapping
        final Map<String, String> mapping = new HashMap<String, String>();

        // Fields
        Field[] fields = excelBeanType.getDeclaredFields();
        for (Field f : fields) {
            String fieldName = f.getName();
            mapping.put(fieldName, fieldName);

            ExcelColumn ec = f.getAnnotation(ExcelColumn.class);
            if (ec != null && StringUtils.isNotEmpty(ec.value())) {
                mapping.put(fieldName, ec.value());
            }
        }

        // Methods
        Method[] methods = excelBeanType.getDeclaredMethods();
        for (Method m : methods) {
            String fieldName = getFieldName(m);
            if (!mapping.containsKey(fieldName)) {
                mapping.put(fieldName, fieldName);
            }

            ExcelColumn ec = m.getAnnotation(ExcelColumn.class);
            if (ec != null && StringUtils.isNotEmpty(ec.value())) {
                mapping.put(fieldName, ec.value());
            }
        }

        LOGGER.info("Bean property to Excel Column of - {} : {}", excelBeanType, mapping);
        return Collections.unmodifiableMap(mapping);
    }

    public static Map<String, String> getColumnToPropertyMap(Class<?> excelBeanType) {
        // Column to Property Mapping
        final Map<String, String> columnToPropMap = new HashMap<String, String>();

        // Bean Property to Column Mapping
        final Map<String, String> propToColumnMap = getPropertyToColumnNameMap(excelBeanType);
        for (String prop : propToColumnMap.keySet()) {
            columnToPropMap.put(propToColumnMap.get(prop), prop);
        }

        LOGGER.info("Excel Column to property map of - {} : {}", excelBeanType, columnToPropMap);
        return Collections.unmodifiableMap(columnToPropMap);
    }

    public static List<String> getColumnNames(Class<?> excelBeanType) {
        // Bean Property to Column Mapping
        final Map<String, String> propToColumnMap = getPropertyToColumnNameMap(excelBeanType);

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
