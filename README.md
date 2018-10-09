[![Build Status](https://travis-ci.org/millij/poi-object-mapper.svg?branch=master)](https://travis-ci.org/millij/poi-object-mapper)
[![codecov](https://codecov.io/gh/millij/poi-object-mapper/branch/master/graph/badge.svg)](https://codecov.io/gh/millij/poi-object-mapper)


# poi-object-mapper

**poi-object-mapper** is a wrapper java library for [Apache POI](https://poi.apache.org/) (Apache POI provides java API to read Microsoft Office Formats). POI APIs are very low level giving acess to all the internals of the file formats.

The aim of this project is to provide easy to use highlevel APIs to read the Office file formats by wrapping the POI APIs. In simple terms, the wrapper APIs would look similar to the [Jackson Project for XML and JSON](https://github.com/FasterXML/jackson), where the data can be mapped to a JAVA Bean and through the mapper APIs, the file data can directly be read as java objects.

*- Note that the current version of the library supports only  **spreadsheets**  (Excel files).*


## Include

This library is not yet available in the Maven Central. For now it hast to be installed manually. Please check the [releases](https://github.com/millij/poi-object-mapper/releases) page for change log and versions.

#### Dependencies

The current implementation uses **POI version 3.17**.


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

Once a mapped Java Bean is ready, use a `Reader` to read the file rows as objects. Use `XlsReader` for `.xls` files and `XlsxReader` for `.xlsx` files.

Reading spreadsheet rows as objects ..

```java
    ...
    final File xlsxFile = new File("<path_to_file>");
    final XlsReader reader = new XlsReader();
    List<Employee> employees = reader.read(xlsxFile, Employee.class);
    ...
```

##### Writing a collection of objects to file

*Currently writing to `.xlsx` files only is supported*

```java
    ...
    // Employees
    List<Employee> employees = new ArrayList<Employee>();
    employees.add(new Employee("1", "foo", 12, "MALE", 1.68));
    employees.add(new Employee("2", "bar", null, "MALE", 1.68));
    employees.add(new Employee("3", "foo bar", null, null, null));
    
    // Writer
    SpreadsheetWriter writer = new SpreadsheetWriter("<output_file_path>");
    writer.addSheet(Employee.class, employees);
    writer.write();
    ...
```

## Implementation Details



## Known Issues

- Reading Large files is not suggested as stream reading is yet to be supported.
- XLS file writing is not supported.
- Mapping `Date` files through a `DateFormat` is not supported yet.
- Reading `Formula` cells is not supported yet.

