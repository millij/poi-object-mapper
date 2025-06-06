package io.github.millij.poi.ss.model.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import io.github.millij.poi.ss.model.DateTimeType;


/**
 * Marker annotation that can be used to define a non-static method as a "setter" or "getter" for a column, or
 * non-static field to be used as a column.
 * 
 * <p>
 * Default value ("") indicates that the field name is used as the column name without any modifications, but it can be
 * specified to non-empty value to specify different name.
 * </p>
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, ElementType.METHOD})
public @interface SheetColumn {

    /**
     * Name of the column to map the annotated property with.
     * 
     * @return column name/header.
     */
    String value() default "";

    /**
     * Setting this to <code>false</code> will enable the null check on the Column values, to ensure non-null values for
     * the field.
     * 
     * default is <code>true</code>. i.e., null values are allowed.
     * 
     * @return <code>true</code> if the annotated field is allowed <code>null</code> as value.
     */
    boolean nullable() default true;


    /**
     * Data presentation format of the Data Column.
     * 
     * @return Data format String.
     */
    String format() default "";

    /**
     * DateTimeType value of the Column
     * 
     * @return the {@link DateTimeType} value of the Column
     */
    DateTimeType datetime() default DateTimeType.NONE;


    // For Write

    /**
     * Column sorting Order when writing content to file
     * 
     * @return Order defined by an integer
     */
    int order() default 0;

}
