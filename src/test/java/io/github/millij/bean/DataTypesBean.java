package io.github.millij.bean;

import static io.github.millij.poi.ss.model.DateTimeType.DATE;
import static io.github.millij.poi.ss.model.DateTimeType.DURATION;

import java.util.Date;

import io.github.millij.poi.ss.model.annotations.Sheet;
import io.github.millij.poi.ss.model.annotations.SheetColumn;


@Sheet
public class DataTypesBean {

    @SheetColumn("ID")
    private String id;

    @SheetColumn("Name")
    private String name;


    // Enum

    @SheetColumn("Gender")
    private Gender gender;


    // Numerical

    @SheetColumn("Age")
    private Integer age;

    @SheetColumn("Height")
    private Double height;


    // DateTime fields

    @SheetColumn(value = "Date", datetime = DATE, format = "dd-MM-yyy")
    private Date date;

    @SheetColumn(value = "Timestamp", datetime = DATE, format = "dd-MM-yyy HH:mm")
    private Long timestamp; // Timestamp

    @SheetColumn(value = "Duration", datetime = DURATION, format = "HH:mm:ss")
    private Long duration;


    // Boolean

    @SheetColumn("Is Active")
    private Boolean isActive;

    @SheetColumn("Is Complete")
    private Boolean isComplete;


    // Constructors
    // ------------------------------------------------------------------------

    public DataTypesBean() {
        // Default
    }


    // Getters and Setters
    // ------------------------------------------------------------------------

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Gender getGender() {
        return gender;
    }

    public void setGender(Gender gender) {
        this.gender = gender;
    }

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public Double getHeight() {
        return height;
    }

    public void setHeight(Double height) {
        this.height = height;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public Long getTimestamp() {
        return timestamp;
    }

    public void setTimestamp(Long timestamp) {
        this.timestamp = timestamp;
    }

    public Long getDuration() {
        return duration;
    }

    public void setDuration(Long duration) {
        this.duration = duration;
    }

    public Boolean getIsActive() {
        return isActive;
    }

    public void setIsActive(Boolean isActive) {
        this.isActive = isActive;
    }

    public Boolean getIsComplete() {
        return isComplete;
    }

    public void setIsComplete(Boolean isComplete) {
        this.isComplete = isComplete;
    }


    // Object Methods
    // ------------------------------------------------------------------------

    @Override
    public String toString() {
        return "DataTypesBean [id=" + id + ", name=" + name + ", gender=" + gender + ", age=" + age + ", height="
                + height + ", date=" + date + ", timestamp=" + timestamp + ", duration=" + duration + ", isActive="
                + isActive + ", isComplete=" + isComplete + "]";
    }

}
