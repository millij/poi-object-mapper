package io.github.millij.poi.ss.writer;

import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import io.github.millij.bean.Company;
import io.github.millij.bean.Employee;


public class XlsxWriterTest {

    private static final Logger LOGGER = LoggerFactory.getLogger(XlsxWriterTest.class);

    private final String _path_test_output = "test-cases/output/";

    // Setup
    // ------------------------------------------------------------------------

    @Before
    public void setup() throws ParseException {
        // prepare
        final File output_dir = new File(_path_test_output);
        if (!output_dir.exists()) {
            output_dir.mkdirs();
        }
    }

    @After
    public void teardown() {
        // nothing to do
    }


    // Tests
    // ------------------------------------------------------------------------

    @Test
    public void test_write_xlsx_single_sheet() throws IOException {
        final String filepath_output_file = _path_test_output.concat("single_sheet.xlsx");

        // Excel Writer
        LOGGER.info("test_write_xlsx_single_sheet :: Writing to file - {}", filepath_output_file);
        SpreadsheetWriter gew = new XlsxWriter();

        // Employees
        List<Employee> employees = new ArrayList<>();
        employees.add(new Employee("1", "foo", 12, "MALE", 1.68));
        employees.add(new Employee("2", "bar", null, "MALE", 1.68));
        employees.add(new Employee("3", "foo bar", null, null, null));

        // Write
        gew.addSheet(Employee.class, employees);
        gew.write(filepath_output_file);
    }

    @Test
    public void test_write_xlsx_single_sheet_custom_headers() throws IOException {
        final String filepath_output_file = _path_test_output.concat("single_sheet_custom_headers.xlsx");

        // Excel Writer
        LOGGER.info("test_write_xlsx_single_sheet :: Writing to file - {}", filepath_output_file);
        SpreadsheetWriter gew = new XlsxWriter();

        // Employees
        List<Employee> employees = new ArrayList<>();
        employees.add(new Employee("1", "foo", 12, "MALE", 1.68));
        employees.add(new Employee("2", "bar", null, "MALE", 1.68));
        employees.add(new Employee("3", "foo bar", null, null, null));

        List<String> headers = Arrays.asList("ID", "Age", "Name", "Address");

        // Add Sheets
        gew.addSheet(Employee.class, employees, headers);

        // Write
        gew.write(filepath_output_file);
    }


    @Test
    public void test_write_xlsx_multiple_sheets() throws IOException {
        final String filepath_output_file = _path_test_output.concat("multiple_sheets.xlsx");

        // Excel Writer
        LOGGER.info("test_write_xlsx_single_sheet :: Writing to file - {}", filepath_output_file);
        SpreadsheetWriter gew = new XlsxWriter();

        // Employees
        List<Employee> employees = new ArrayList<>();
        employees.add(new Employee("1", "foo", 12, "MALE", 1.68));
        employees.add(new Employee("2", "bar", null, "MALE", 1.68));
        employees.add(new Employee("3", "foo bar", null, null, null));

        // Companies
        List<Company> companies = new ArrayList<>();
        companies.add(new Company("Google", 12000, "Palo Alto, CA"));
        companies.add(new Company("Facebook", null, "Mountain View, CA"));
        companies.add(new Company("SpaceX", null, null));

        // Add Sheets
        gew.addSheet(Employee.class, employees);
        gew.addSheet(Company.class, companies);

        // Write
        gew.write(filepath_output_file);
    }


    //
    // Write from Map

    @Test
    public void test_write_xlsx_from_map() throws IOException {
        final String filepath_output_file = _path_test_output.concat("single_sheet_map_data.xlsx");

        // Excel Writer
        LOGGER.info("#test_write_xlsx_from_map :: Writing to file - {}", filepath_output_file);
        final SpreadsheetWriter gew = new XlsWriter();

        // Headers
        final List<String> headers = new ArrayList<>();
        headers.add("S.No.");
        headers.add("Name");
        headers.add("Age");
        headers.add("Gender");
        headers.add("Height (mts)");
        headers.add("Address");
        headers.add("DOB");
        headers.add("Is Alive ?");

        // Data
        final Map<String, Object> row1 = new HashMap<>();
        row1.put("S.No.", 1);
        row1.put("Name", "foo");
        row1.put("Age", 1);
        row1.put("Gender", null);
        row1.put("Height (mts)", "");
        row1.put("Address", "Chennai, India");
        row1.put("Is Alive ?", Boolean.TRUE);

        final Map<String, Object> row2 = new HashMap<>();
        row2.put("S.No.", 4);
        row2.put("Name", "John");
        row2.put("Height (mts)", 1.6D);
        row2.put("DOB", new Date());
        row2.put("Unknown_Header", "unknown");
        row2.put("Is Alive ?", Boolean.FALSE);

        final List<Map<String, Object>> rowsData = Arrays.asList(row1, null, new HashMap<>(), row2);

        // Add Sheets
        final String sheetName = "test_sheet";
        gew.addSheet(rowsData, sheetName, headers);

        // Write
        gew.write(filepath_output_file);
    }

    @Test
    public void test_write_xlsx_from_map_default_sheetname() throws IOException {
        final String filepath_output_file = _path_test_output.concat("single_sheet_map_data_default_sheetname.xlsx");

        // Excel Writer
        LOGGER.info("#test_write_xlsx_from_map_default_sheetname :: Writing to file - {}", filepath_output_file);
        final SpreadsheetWriter gew = new XlsWriter();

        // Headers
        final List<String> headers = new ArrayList<>();
        headers.add("S.No.");
        headers.add("Name");
        headers.add("Age");
        headers.add("Gender");
        headers.add("Height (mts)");
        headers.add("Address");

        // Data
        final Map<String, Object> row1 = new HashMap<>();
        row1.put("S.No.", 1);
        row1.put("Name", "foo");
        row1.put("Age", 1);
        row1.put("Gender", "Male");
        row1.put("Unknown_Header", null);

        final List<Map<String, Object>> rowsData = Arrays.asList(row1);

        // Add Sheets
        gew.addSheet(rowsData, null, headers);

        // Write
        final File outFile = new File(filepath_output_file);
        gew.write(outFile);
    }


}
