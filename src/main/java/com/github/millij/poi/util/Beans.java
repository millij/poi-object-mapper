package com.github.millij.poi.util;

import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.lang.reflect.Method;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public final class Beans {

    private static final Logger LOGGER = LoggerFactory.getLogger(Spreadsheet.class);

    private Beans() {
        // Utility Class
    }


    // Static Utilities
    // ------------------------------------------------------------------------

    public static String getFieldName(Method method) {
        // Sanity checks
        if (method == null) {
            return null;
        }

        String methodName = method.getName();
        return Introspector.decapitalize(methodName.substring(methodName.startsWith("is") ? 2 : 3));
    }

    public static String getFieldValueAsString(Object beanObj, String fieldName) throws Exception {
        // Sanity checks
        PropertyDescriptor pd = new PropertyDescriptor(fieldName, beanObj.getClass());
        Method getterMtd = pd.getReadMethod();

        Object value = getterMtd.invoke(beanObj);
        String cellValue = value != null ? value.toString() : "";

        return cellValue;
    }



}
