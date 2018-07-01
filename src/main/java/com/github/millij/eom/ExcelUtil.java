package com.github.millij.eom;

import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.File;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
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

import com.github.millij.eom.spi.IExcelBean;
import com.github.millij.eom.spi.annotation.ExcelColumn;


public class ExcelUtil {

    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelUtil.class);

    private ExcelUtil() {
        // Utility Class
    }

    // Labels
    // ------------------------------------------------------------------------

    public final static String EXTN_XLS = "xls";
    public final static String EXTN_XLSX = "xlsx";


    // Utilities
    // ------------------------------------------------------------------------

    // File

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

    public static Map<String, String> asRowDataMap(IExcelBean beanObj, List<String> headers)
            throws IllegalAccessException, InvocationTargetException, IntrospectionException {
        // Sanity checks
        if (beanObj == null || CollectionUtils.isEmpty(headers)) {
            return new HashMap<>();
        }


        // RowData map
        final Map<String, String> rowDataMap = new HashMap<String, String>();

        for (Field f : beanObj.getClass().getDeclaredFields()) {
            if (!f.isAnnotationPresent(ExcelColumn.class)) {
                continue;
            }

            ExcelColumn excelColumn = f.getAnnotation(ExcelColumn.class);
            String header = excelColumn.value();
            if (!headers.contains(header)) {
                continue;
            }

            PropertyDescriptor pd = new PropertyDescriptor(f.getName(), beanObj.getClass());
            Method getterMtd = pd.getReadMethod();

            Object value = getterMtd.invoke(beanObj);
            String cellValue = value != null ? value.toString() : "";

            rowDataMap.put(header, cellValue);
        }

        return rowDataMap;
    }


 
    // Bean :: Property Utils

    public static Map<String, String> getPropertyToColumnNameMap(Class<? extends IExcelBean> excelBeanType) {
        // Sanity checks
        if (excelBeanType == null) {
            throw new IllegalArgumentException("getColumnToPropertyMap :: Invalid ExcelBean type - " + excelBeanType);
        }

        // Property to Column name Mapping
        final Map<String, String> mapping = new HashMap<String, String>();

        // Check fields
        Field[] fields = excelBeanType.getDeclaredFields();
        for (Field f : fields) {
            String fieldName = f.getName();
            mapping.put(fieldName, fieldName);

            Annotation annotation = f.getAnnotation(ExcelColumn.class);
            ExcelColumn ec = (ExcelColumn) annotation;
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

            Annotation annotation = m.getAnnotation(ExcelColumn.class);
            ExcelColumn ec = (ExcelColumn) annotation;
            if (ec != null && StringUtils.isNotEmpty(ec.value())) {
                mapping.put(fieldName, ec.value());
            } 
        }

        LOGGER.info("Bean property to Excel Column of - {} : {}", excelBeanType, mapping);
        return Collections.unmodifiableMap(mapping);
    }

    public static Map<String, String> getColumnToPropertyMap(Class<? extends IExcelBean> excelBeanType) {
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

    public static List<String> getColumnNames(Class<? extends IExcelBean> excelBeanType) {
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



}
