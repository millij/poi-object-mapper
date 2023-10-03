package io.github.millij.bean;

import io.github.millij.poi.ss.model.annotations.Sheet;
import io.github.millij.poi.ss.model.annotations.SheetColumn;


@Sheet
public class Employee {

    // Note that Id and Name are annotated at name level
    private String id;
    private String name;

    @SheetColumn("Age")
    private Integer age;

    @SheetColumn("Gender")
    private String gender;

    @SheetColumn("Height (mts)")
    private Double height;

    @SheetColumn("Address")
    private String address;

    @SheetColumn(value = "Formula", isFormula = true)
    private String formula;

    @SheetColumn(value = "Date", isFormatted = true, format = "yyyy/MM/dd")
    private String date;

    // Constructors
    // ------------------------------------------------------------------------

    public Employee() {
        // Default
    }

    public Employee(String id, String name, Integer age, String gender, Double height, String formula, String date,
            String address) {
        super();

        this.id = id;
        this.name = name;
        this.age = age;
        this.gender = gender;
        this.height = height;
        this.date = date;
        this.address = address;
        this.formula = formula;
    }

    // Getters and Setters
    // ------------------------------------------------------------------------

    @SheetColumn("ID")
    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    @SheetColumn("Name")
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

    public String getDate() {
        return date;
    }

    public void setDate(String date) {
        this.date = date;
    }

    public String getFormula() {
        return formula;
    }

    public void setFormula(String formula) {
        this.formula = formula;
    }

    // Object Methods
    // ------------------------------------------------------------------------

    @Override
    public String toString() {
        return "Employee [id=" + id + ", name=" + name + ", age=" + age + ", gender=" + gender + ", height=" + height
                + ", date=" + date + ", Address =" + address + ", Formula=" + formula + "]";
    }

}
