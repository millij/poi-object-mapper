package io.github.millij.bean;

import java.time.LocalDate;
import java.util.Date;

// import java.util.Date;

import io.github.millij.poi.ss.model.annotations.Sheet;
import io.github.millij.poi.ss.model.annotations.SheetColumn;


@Sheet("Schedules")
public class Schedule {

    @SheetColumn("Day")
    private String day;

    @SheetColumn(value = "Date", format = "dd/MM/yyyy")
    private Date date;

    @SheetColumn(value = "localDate", format = "dd/MM/yyyy")
    private LocalDate localDate;

    // Constructor
    // -------------------------------------------------------------------------

    public Schedule(String day, Date date, LocalDate localDate) {
        super();
        this.day = day;
        this.date = date;
        this.localDate = localDate;
    }

    // Getters and Setters
    // -------------------------------------------------------------------------

    public String getDay() {
        return day;
    }

    public void setDay(String day) {
        this.day = day;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public LocalDate getLocalDate() {
        return localDate;
    }

    public void setLocalDate(LocalDate localDate) {
        this.localDate = localDate;
    }

    // Object Methods
    // ------------------------------------------------------------------------
    @Override
    public String toString() {
        return "Schedule [day=" + day + ", date=" + date + ", localDate=" + localDate + "]";
    }

}
