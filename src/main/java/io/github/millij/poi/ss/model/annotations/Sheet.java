package io.github.millij.poi.ss.model.annotations;

import static java.lang.annotation.ElementType.TYPE;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;


/**
 * Marker annotation that can be used to define a "type" for a sheet. The value of this annotation will be used to map
 * the sheet of the workbook to this bean definition.
 * 
 * <p>
 * Default value ("") indicates that the default sheet name to be used without any modifications, but it can be
 * specified to non-empty value to specify different name.
 * </p>
 * 
 * <p>
 * Typically used when writing the java objects to the file.
 * </p>
 */
@Retention(RUNTIME)
@Target(TYPE)
public @interface Sheet {

    /**
     * Name of the sheet.
     * 
     * @return the Sheet name
     */
    String value() default "";

}
