package com.github.millij.eom.bean;

import com.github.millij.eom.spi.IExcelBean;
import com.github.millij.eom.spi.annotation.ExcelColumn;
import com.github.millij.eom.spi.annotation.ExcelSheet;


@ExcelSheet("Employees")
public class Employee implements IExcelBean {

    // Note that Id and Name are annotated at name level
    private String id;
    private String name;

    @ExcelColumn("Age")
    private Integer age;

    @ExcelColumn("Gender")
    private String gender;

    @ExcelColumn("Height (mts)")
    private Double height;

    @ExcelColumn("Address")
    private String address;


    // Constructors
    // ------------------------------------------------------------------------

    public Employee() {
        // Default
    }

    public Employee(String id, String name, Integer age, String gender, Double height) {
        super();

        this.id = id;
        this.name = name;
        this.age = age;
        this.gender = gender;
        this.height = height;
    }


    // Getters and Setters
    // ------------------------------------------------------------------------

    @ExcelColumn("ID")
    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    @ExcelColumn("Name")
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
