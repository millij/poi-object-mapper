package io.github.millij.poi.ss.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import io.github.millij.bean.Company;
import io.github.millij.bean.Employee;
import io.github.millij.poi.SpreadsheetReadException;
import io.github.millij.poi.ss.handler.RowListener;


public class XlsReaderTest {

    private static final Logger LOGGER = LoggerFactory.getLogger(XlsReaderTest.class);

    // XLS
    private String _filepath_xls_single_sheet;
    private String _filepath_xls_multiple_sheets;

    // Setup
    // ------------------------------------------------------------------------

    @Before
    public void setup() throws ParseException {
        // filepaths

        // xls
        _filepath_xls_single_sheet = "src/test/resources/sample-files/xls_sample_single_sheet.xls";
        _filepath_xls_multiple_sheets = "src/test/resources/sample-files/xls_sample_multiple_sheets.xls";
    }

    @After
    public void teardown() {
        // nothing to do
    }


    // Tests
    // ------------------------------------------------------------------------


    // Read from file

    @Test
    public void test_read_xls_single_sheet() throws SpreadsheetReadException {
        // Excel Reader
        LOGGER.info("test_read_xls_single_sheet :: Reading file - {}", _filepath_xls_single_sheet);
        XlsReader reader = new XlsReader();

        // Read
        List<Employee> employees = reader.read(Employee.class, new File(_filepath_xls_single_sheet));
        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);

        for (Employee emp : employees) {
            LOGGER.info("test_read_xls_single_sheet :: Output - {}", emp);
        }
    }

    @Test
    public void test_read_xls_multiple_sheets() throws SpreadsheetReadException {
        // Excel Reader
        LOGGER.info("test_read_xls_multiple_sheets :: Reading file - {}", _filepath_xls_multiple_sheets);
        XlsReader reader = new XlsReader();

        // Read Sheet 1
        List<Employee> employees = reader.read(Employee.class, new File(_filepath_xls_multiple_sheets), 1);
        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);

        for (Employee emp : employees) {
            LOGGER.info("test_read_xls_multiple_sheets :: Output - {}", emp);
        }

        // Read Sheet 2
        List<Company> companies = reader.read(Company.class, new File(_filepath_xls_multiple_sheets), 2);
        Assert.assertNotNull(companies);
        Assert.assertTrue(companies.size() > 0);

        for (Company company : companies) {
            LOGGER.info("test_read_xls_multiple_sheets :: Output - {}", company);
        }
    }


    // Read from Stream

    @Test
    public void test_read_xls_single_sheet_from_stream() throws SpreadsheetReadException, FileNotFoundException {
        // Excel Reader
        LOGGER.info("test_read_xls_single_sheet_from_stream :: Reading file - {}", _filepath_xls_single_sheet);
        XlsReader reader = new XlsReader();

        // InputStream
        final InputStream fis = new FileInputStream(new File(_filepath_xls_single_sheet));

        // Read
        List<Employee> employees = reader.read(Employee.class, fis);
        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);

        for (Employee emp : employees) {
            LOGGER.info("test_read_xls_single_sheet :: Output - {}", emp);
        }
    }

    @Test
    public void test_read_xls_multiple_sheets_from_stream() throws SpreadsheetReadException, FileNotFoundException {
        // Excel Reader
        LOGGER.info("test_read_xls_multiple_sheets_from_stream :: Reading file - {}", _filepath_xls_multiple_sheets);
        XlsReader reader = new XlsReader();

        // InputStream
        final InputStream fisSheet1 = new FileInputStream(new File(_filepath_xls_multiple_sheets));

        // Read Sheet 1
        List<Employee> employees = reader.read(Employee.class, fisSheet1, 1);
        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);

        for (Employee emp : employees) {
            LOGGER.info("test_read_xls_multiple_sheets :: Output - {}", emp);
        }

        // InputStream
        final InputStream fisSheet2 = new FileInputStream(new File(_filepath_xls_multiple_sheets));

        // Read Sheet 2
        List<Company> companies = reader.read(Company.class, fisSheet2, 2);
        Assert.assertNotNull(companies);
        Assert.assertTrue(companies.size() > 0);

        for (Company company : companies) {
            LOGGER.info("test_read_xls_multiple_sheets :: Output - {}", company);
        }
    }


    // Read with Callback

    @Test
    public void test_read_xls_single_sheet_with_callback() throws SpreadsheetReadException {
        // Excel Reader
        LOGGER.info("test_read_xls_single_sheet_with_callback :: Reading file - {}", _filepath_xls_single_sheet);

        // file
        final File xlsFile = new File(_filepath_xls_single_sheet);

        final List<Employee> employees = new ArrayList<Employee>();

        // Read
        XlsReader reader = new XlsReader();
        reader.read(Employee.class, xlsFile, new RowListener<>() {

            @Override
            public void row(int rowNum, Employee employee) {
                employees.add(employee);
                LOGGER.info("test_read_xls_single_sheet_with_callback :: Output - {}", employee);

            }

        });

        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);
    }

}
