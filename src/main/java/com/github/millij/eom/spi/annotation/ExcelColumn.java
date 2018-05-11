package com.github.millij.eom.spi.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, ElementType.METHOD})
public @interface ExcelColumn {

    /**
     * Name of the column to map the annotated property with.
     */
    String value();

    /**
     * Setting this to <code>false</code> will enable the null check on the Column values, to ensure
     * non-null values for the field.
     * 
     * default is <code>true</code>. i.e., null values are allowed.
     */
    boolean nullable() default true;

}
