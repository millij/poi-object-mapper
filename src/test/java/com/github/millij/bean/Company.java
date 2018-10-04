package com.github.millij.bean;

import com.github.millij.poi.ss.model.SheetColumn;
import com.github.millij.poi.ss.model.Sheet;


@Sheet("Companies")
public class Company {

    @SheetColumn("Company Name")
    private String name;

    @SheetColumn("# of Employees")
    private Integer noOfEmployees;

    @SheetColumn("Address")
    private String address;


    // Constructors
    // ------------------------------------------------------------------------

    public Company() {
        // Default
    }

    public Company(String name, Integer noOfEmployees, String address) {
        super();

        this.name = name;
        this.noOfEmployees = noOfEmployees;
        this.address = address;
    }


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
