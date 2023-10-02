package io.github.millij.poi.ss.model.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;


/**
 * Marker annotation that can be used to define a non-static method as a "setter" or "getter" for a
 * column, or non-static field to be used as a column.
 * 
 * <p>
 * Default value ("") indicates that the field name is used as the column name without any
 * modifications, but it can be specified to non-empty value to specify different name.
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
    boolean isFormatted() default false;
    String format() default "dd/MM/yyyy";
    
    boolean isFromula() default false;

    /**
     * Setting this to <code>false</code> will enable the null check on the Column values, to ensure
     * non-null values for the field.
     * 
     * default is <code>true</code>. i.e., null values are allowed.
     * 
     * @return <code>true</code> if the annotated field is allowed <code>null</code> as value.
     */
    boolean nullable() default true;

}
