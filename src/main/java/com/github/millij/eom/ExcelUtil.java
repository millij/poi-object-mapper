package com.github.millij.eom;

import java.io.File;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.millij.eom.spi.IExcelEntity;
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


    public static Map<String, String> getColumnToPropertyMap(Class<? extends IExcelEntity> entityType) {
        // Sanity checks
        if (entityType == null) {
            throw new IllegalArgumentException("getColumnToPropertyMap :: Invalid entity type - " + entityType);
        }

        // Column to Property Mapping
        Map<String, String> mapping = new HashMap<String, String>();

        Field[] fields = entityType.getDeclaredFields();
        if (fields == null || fields.length == 0) {
            return mapping;
        }

        for (int i = 0; i < fields.length; i++) {
            Field f = fields[i];

            Annotation annotation = f.getAnnotation(ExcelColumn.class);
            ExcelColumn ec = (ExcelColumn) annotation;

            if (ec != null && ec.value() != null && !ec.value().isEmpty()) {
                mapping.put(ec.value(), f.getName());
            }
        }

        return Collections.unmodifiableMap(mapping);
    }


    public static String getCellColumnReference(String cellRef) {
        // Sanity checks
        if (StringUtils.isEmpty(cellRef)) {
            return "";
        }

        // Splits the Cell name and returns the column reference
        String cellColRef = cellRef.split("[0-9]*$")[0];
        return cellColRef;
    }

}
