package io.github.millij.bean;

import io.github.millij.poi.ss.model.annotations.SheetColumn;

public class Department {

    @SheetColumn("Company Name")
    private String name;

    @SheetColumn("# of Employees")
    private Integer noOfEmployees;

    // Constructors
    // ------------------------------------------------------------------------

    public Department() {
	// Default
    }

    public Department(String name, Integer noOfEmployees) {
	super();

	this.name = name;
	this.noOfEmployees = noOfEmployees;
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
    
    // Object Methods
    // ------------------------------------------------------------------------

    @Override
    public String toString() {
        return "Department [name=" + name + ", noOfEmployees=" + noOfEmployees + "]";
    }
}
