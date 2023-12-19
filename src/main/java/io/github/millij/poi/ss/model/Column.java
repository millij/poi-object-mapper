package io.github.millij.poi.ss.model;

import java.util.Objects;

import io.github.millij.poi.ss.model.annotations.SheetColumn;


/**
 * Bean representation of Spreadsheet column definition.
 * 
 * @see SheetColumn
 */
public class Column implements Comparable<Column> {

    private String name;
    private Boolean nullable;
    private String format;

    private Integer order;

    private DateTimeType datetimeType;


    // Constructors
    // ------------------------------------------------------------------------

    public Column() {
        super();
    }

    public Column(String name) {
        super();

        // init
        this.name = name;
    }


    // Methods
    // ------------------------------------------------------------------------

    // Comparable

    @Override
    public int compareTo(final Column o) {
        if (Objects.isNull(o) || Objects.isNull(o.getOrder())) {
            return 1;
        }

        return Objects.isNull(this.getOrder()) ? -1 : Integer.compare(this.getOrder(), o.getOrder());
    }


    // Getters and Setters
    // ------------------------------------------------------------------------

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Boolean getNullable() {
        return nullable;
    }

    public void setNullable(Boolean nullable) {
        this.nullable = nullable;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public Integer getOrder() {
        return order;
    }

    public void setOrder(Integer order) {
        this.order = order;
    }

    public DateTimeType getDatetimeType() {
        return datetimeType;
    }

    public void setDatetimeType(DateTimeType datetimeType) {
        this.datetimeType = datetimeType;
    }


    // Object Methods
    // ------------------------------------------------------------------------

    @Override
    public String toString() {
        return "Column [name=" + name + ", nullable=" + nullable + ", format=" + format + ", order=" + order
                + ", datetimeType=" + datetimeType + "]";
    }


}
