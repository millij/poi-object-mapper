
[![Build Status](https://travis-ci.org/millij/poi-object-mapper.svg?branch=master)](https://travis-ci.org/millij/poi-object-mapper)
[![codecov](https://codecov.io/gh/millij/poi-object-mapper/branch/master/graph/badge.svg)](https://codecov.io/gh/millij/poi-object-mapper)


# poi-object-mapper

**poi-object-mapper** is a wrapper java library for [Apache POI](https://poi.apache.org/) (Apache POI provides java API to read Microsoft Office Formats). POI APIs are very low level giving access to all the internals of the file formats.

The aim of this project is to provide easy to use high-level APIs to read the Office file formats by wrapping the POI APIs. In simple terms, the wrapper APIs would look similar to the [Jackson Project for XML and JSON](https://github.com/FasterXML/jackson), where the data can be mapped to a JAVA Bean and through the mapper APIs, the file data can directly be read as java objects.

*- Note that the current version of the library supports only  **spreadsheets**  (Excel files).*


## Include

This library is available in [Maven Central](https://mvnrepository.com/artifact/io.github.millij/poi-object-mapper). 

`pom.xml` entry details..

```
<dependency>
    <groupId>io.github.millij</groupId>
    <artifactId>poi-object-mapper</artifactId>
    <version>3.2.0</version>
</dependency>
```

To install manually, please check the [releases](https://github.com/millij/poi-object-mapper/releases) page for available versions and  change log.

#### Dependencies

The current implementation uses **POI version 5.2.5**.


## Usage

### Spreadsheets (Excel)

Consider the below sample spreadsheet, where data of employees is present.

| Name              | Age   | Gender | Height (mts) | Address                            |
| ----------------- |:----- | :----- | ------------:| :--------------------------------- |
| Bob               | 32    | MALE   | 1.8          | 410, Madison, Seattle, WA – 123456 |
| John Doe          | 45    | MALE   | 2.1          |                                    |
| Guiliano Carlini  |       | MALE   | 1.78         | Palo Alto, CA – 43234              |


##### Mapping Rows to a Java Bean

Create a java bean and map its properties to the columns using the `@SheetColumn` annotation. The `@SheetColumn` annotation can be declared on the `Field`, as well as its `Accessor Methods`. Pick any one of them to configure the mapped `Column` as per convenience.

```java
@Sheet
public class Employee {
    // Pick either field or its accessor methods to apply the Column mapping.
    ...
    @SheetColumn("Age")
    private Integer age;
    ...
    @SheetColumn("Name")
    public String getName() {
        return name;
    }
    ...
}
```

##### Reading Rows as Java Objects

Once a mapped Java Bean is ready, use a `Reader` to read the file rows as objects.

Use `XlsReader` for `.xls` files and `XlsxReader` for `.xlsx` files.

Reading spreadsheet rows as objects ..

```java
    ...
    final File xlsFile = new File("<path_to_file>");
    final XlsReader reader = new XlsReader();
    final List<Employee> employees = reader.read(Employee.class, xlsFile);
    ...
```

##### Reading Rows as Map (when there is no mapping bean)

Reading spreadsheet rows as `Map<String, Object>` Objects ..

```java
    ...
    final File xlsxFile = new File("<path_to_file>");
    final XlsxReader reader = new XlsxReader(); // OR XlsReader as needed
    final List<Map<String, Object>> rowObjects = reader.read(xlsxFile);
    ...
```


##### Writing a List of objects to file

Similar to `Reader`, the mapped Java Beans can be written to files.

Use `XlsWriter` for `.xls` files and `XlsxWriter` for `.xlsx` files.

```java
    ...
    // Employees
    final List<Employee> employees = new ArrayList<>();
    employees.add(new Employee("1", "foo", 12, "MALE", 1.68));
    employees.add(new Employee("2", "bar", null, "MALE", 1.68));
    employees.add(new Employee("3", "foo bar", null, null, null));

    // Writer
    final SpreadsheetWriter writer = new XlsxWriter();
    writer.addSheet(Employee.class, employees);
    writer.write("<output_file_path>");
    ...
```

##### Writing a List of objects defined as Map to file

Similar to `Reader`, the `Writer` also supports Row data defined as `Map<String, Object>`. 
This is to support the data that is not backed by a concrete bean Class definition

Use `XlsWriter` for `.xls` files and `XlsxWriter` for `.xlsx` files.

```java
    ...
    // Employees
    final Map<String, Object> emp1 = new HashMap<>();
    emp1.put("Name", "foo");
    emp1.put("Age", 12);
    emp1.put("Gender", "MALE");
    emp1.put("Height (mts)", 1.68);

    final Map<String, Object> emp2 = new HashMap<>();
    emp1.put("Name", "bar");
    emp1.put("Age", null);
    emp1.put("Gender", "MALE");

    final List<Map<String, Object>> employees = Arrays.asList(emp1, emp2);

    // Headers (or Column Names)
    final List<String> headers = new ArrayList<>();
    headers.add("Name");
    headers.add("Age");
    headers.add("Gender");
    headers.add("Height (mts)");
    headers.add("Address");

    // Writer
    final SpreadsheetWriter writer = new XlsxWriter();
    writer.addSheet(employees, headers);
    writer.write("<output_file_path>");

```



## Issues

The known issues are already listed under [Issues Section](https://github.com/millij/poi-object-mapper/issues).

Please add there your bugs findings, feature requests, enhancements etc.




