package com.github.millij.eom.bean;

import com.github.millij.eom.spi.IExcelEntity;
import com.github.millij.eom.spi.annotation.ExcelColumn;


public class Company implements IExcelEntity {

    @ExcelColumn("Company Name")
    private String name;

    @ExcelColumn("# of Employees")
    private Integer noOfEmployees;

    @ExcelColumn("Address")
    private String address;

    
    // Getters and Setters
    // ------------------------------------------------------------------------
    
    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getNoOfEmployees() {
        return noOfEmployees;
    }

    public void setNoOfEmployees(Integer noOfEmployees) {
        this.noOfEmployees = noOfEmployees;
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
        return "Company [name=" + name + ", noOfEmployees=" + noOfEmployees + ", address=" + address + "]";
    }

}
