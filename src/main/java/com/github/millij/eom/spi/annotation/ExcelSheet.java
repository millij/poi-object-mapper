package com.github.millij.eom.spi.annotation;

import static java.lang.annotation.ElementType.TYPE;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;


@Retention(RUNTIME)
@Target(TYPE)
public @interface ExcelSheet {

    /**
     * Name of the sheet.
     */
    String value() default "";

}
