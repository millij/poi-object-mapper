package io.github.millij.poi.ss.model;


/**
 * Time of Date Time value String
 */
public enum DateTimeType {

    /**
     * Complete Date object (with both Date and Time components), i.e., an equivalent of timestamp.
     * 
     * ex. "18-12-2021 19:30" (format: dd-MM-yyyy HH:mm)
     */
    DATE,

    /**
     * Time of the Day (starting from the midnight). Can have AM/PM.
     * 
     * ex. "19:30" (format: HH:mm)
     */
    TIME,

    /**
     * Difference between two timestamps
     * 
     * ex. "06:30" (format: HH:mm)
     */
    DURATION,


    /**
     * None
     */
    NONE;

}
