package com.github.millij.eom.bean;

import com.github.millij.eom.spi.IExcelEntity;
import com.github.millij.eom.spi.annotation.ExcelColumn;



public class Employee implements IExcelEntity {

    @ExcelColumn("ID")
    private String id;

    @ExcelColumn("Name")
    private String name;

    @ExcelColumn("Age")
    private Integer age;

    @ExcelColumn("Gender")
    private String gender;

    @ExcelColumn("Height (mts)")
    private Double height;

    @ExcelColumn("Address")
    private String address;


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

    public Integer getAge() {
        return age;
    }

    public void setAge(Integer age) {
        this.age = age;
    }

    public String getGender() {
        return gender;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public Double getHeight() {
        return height;
    }

    public void setHeight(Double height) {
        this.height = height;
    }

    public String getAddress() {
        return address;
    }

    public void setAddress(String address) {
        this.address = address;
    }


    // Object Methods
    // ------------------------------------------------------------------------

    @Override
    public String toString() {
        return "Employee [id=" + id + ", name=" + name + ", age=" + age + ", gender=" + gender + ", height=" + height
                + ", address=" + address + "]";
    }


}
