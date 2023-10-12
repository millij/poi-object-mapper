package io.github.millij.bean;

import io.github.millij.poi.ss.model.annotations.SheetColumn;

public class Emp_sameCols 
{
    private String id;
    private String name;

    @SheetColumn(value="Age",index=2)
    private Integer age;

    @SheetColumn(value="Gender",index=3)
    private String gender;

    @SheetColumn(value="Height (mts)")
    private Double height;
    
    @SheetColumn(value="Address")
    private String address;
    
    @SheetColumn(value="Name",index=6)
    private String name2;
    
    
    public Emp_sameCols() {
        // Default
    }

    public Emp_sameCols(String id, String name, Integer age, String gender, Double height,String address,String name2) {
        super();

        this.id = id;
        this.name = name;
        this.age = age;
        this.gender = gender;
        this.height = height;
        this.address = address;
        this.name2 = name2;
        
    }
    
    
    @SheetColumn(value="ID",index=0)
    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    @SheetColumn(value="Name",index=1)
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
    
    public String getName2() {
        return name2;
    }

    public void setName2(String name2) {
        this.name2 = name2;
    }

    @Override
    public String toString() {
        return "Emp_indexed [id=" + id + ", name=" + name + ", age=" + age + ", gender=" + gender + ", height=" + height
                + ", address=" + address +", Name2="+name2+"]";
    }
    
    
    
}