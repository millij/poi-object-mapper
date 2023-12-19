package io.github.millij.poi.ss.writer;

import io.github.millij.bean.Company;
import io.github.millij.bean.Employee;
import io.github.millij.poi.ss.writer.SpreadsheetWriter;

import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class SpreadsheetWriterTest {

    private static final Logger LOGGER = LoggerFactory.getLogger(SpreadsheetWriterTest.class);

    private final String _path_test_output = "test-cases/output/";

    // Setup
    // ------------------------------------------------------------------------

    @Before
    public void setup() throws ParseException {
        // prepare
        File output_dir = new File(_path_test_output);
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
        SpreadsheetWriter gew = new SpreadsheetWriter(filepath_output_file);

        // Employees
        List<Employee> employees = new ArrayList<Employee>();
        employees.add(new Employee("1", "foo", 12, "MALE", 1.68));
        employees.add(new Employee("2", "bar", null, "MALE", 1.68));
        employees.add(new Employee("3", "foo bar", null, null, null));

        // Write
        gew.addSheet(Employee.class, employees);
        gew.write();
    }

    @Test
    public void test_write_xlsx_single_sheet_custom_headers() throws IOException {
        final String filepath_output_file = _path_test_output.concat("single_sheet_custom_headers.xlsx");

        // Excel Writer
        LOGGER.info("test_write_xlsx_single_sheet :: Writing to file - {}", filepath_output_file);
        SpreadsheetWriter gew = new SpreadsheetWriter(filepath_output_file);

        // Employees
        List<Employee> employees = new ArrayList<Employee>();
        employees.add(new Employee("1", "foo", 12, "MALE", 1.68));
        employees.add(new Employee("2", "bar", null, "MALE", 1.68));
        employees.add(new Employee("3", "foo bar", null, null, null));

        List<String> headers = Arrays.asList("ID", "Age", "Name", "Address");

        // Add Sheets
        gew.addSheet(Employee.class, employees, headers);

        // Write
        gew.write();
    }


    @Test
    public void test_write_xlsx_multiple_sheets() throws IOException {
        final String filepath_output_file = _path_test_output.concat("multiple_sheets.xlsx");

        // Excel Writer
        LOGGER.info("test_write_xlsx_single_sheet :: Writing to file - {}", filepath_output_file);
        SpreadsheetWriter gew = new SpreadsheetWriter(filepath_output_file);

        // Employees
        List<Employee> employees = new ArrayList<Employee>();
        employees.add(new Employee("1", "foo", 12, "MALE", 1.68));
        employees.add(new Employee("2", "bar", null, "MALE", 1.68));
        employees.add(new Employee("3", "foo bar", null, null, null));

        // Campanies
        List<Company> companies = new ArrayList<Company>();
        companies.add(new Company("Google", 12000, "Palo Alto, CA"));
        companies.add(new Company("Facebook", null, "Mountain View, CA"));
        companies.add(new Company("SpaceX", null, null));

        // Add Sheets
        gew.addSheet(Employee.class, employees);
        gew.addSheet(Company.class, companies);

        // Write
        gew.write();
    }

}
