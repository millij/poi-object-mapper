package io.github.millij.bean;

import io.github.millij.poi.ss.model.annotations.Sheet;
import io.github.millij.poi.ss.model.annotations.SheetColumn;


@Sheet
public class EmployeeIndexed {
    private String id;
    private String name;

    @SheetColumn(value = "Age", index = 2)
    private Integer age;

    @SheetColumn(value = "Gender", index = 3)
    private String gender;

    @SheetColumn(value = "Height (mts)", index = 4)
    private Double height;

    @SheetColumn(value = "Address", index = 5)
    private String address;


    public EmployeeIndexed() {
        // Default
    }

    public EmployeeIndexed(String id, String name, Integer age, String gender, Double height, String address) {
        super();

        this.id = id;
        this.name = name;
        this.age = age;
        this.gender = gender;
        this.height = height;
        this.address = address;

    }



    @SheetColumn(value = "ID", index = 0)
    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    @SheetColumn(value = "Name", index = 1)
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

    @Override
    public String toString() {
        return "Emp_indexed [id=" + id + ", name=" + name + ", age=" + age + ", gender=" + gender + ", height=" + height
                + ", address=" + address + "]";
    }
}
