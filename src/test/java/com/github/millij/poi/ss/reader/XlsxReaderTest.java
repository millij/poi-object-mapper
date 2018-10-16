package com.github.millij.poi.ss.reader;

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

import com.github.millij.bean.Company;
import com.github.millij.bean.Employee;
import com.github.millij.poi.SpreadsheetReadException;
import com.github.millij.poi.ss.handler.RowListener;


public class XlsxReaderTest {

    private static final Logger LOGGER = LoggerFactory.getLogger(XlsxReaderTest.class);

    // XLSX
    private String _filepath_xlsx_single_sheet;
    private String _filepath_xlsx_multiple_sheets;


    // Setup
    // ------------------------------------------------------------------------

    @Before
    public void setup() throws ParseException {
        // sample files
        _filepath_xlsx_single_sheet = "src/test/resources/sample-files/xlsx_sample_single_sheet.xlsx";
        _filepath_xlsx_multiple_sheets = "src/test/resources/sample-files/xlsx_sample_multiple_sheets.xlsx";
    }

    @After
    public void teardown() {
        // nothing to do
    }


    // Tests
    // ------------------------------------------------------------------------


    // Read from file

    @Test
    public void test_read_xlsx_single_sheet() throws SpreadsheetReadException {
        // Excel Reader
        LOGGER.info("test_read_xlsx_single_sheet :: Reading file - {}", _filepath_xlsx_single_sheet);
        XlsxReader reader = new XlsxReader();

        // Read
        List<Employee> employees = reader.read(Employee.class, new File(_filepath_xlsx_single_sheet));
        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);

        for (Employee emp : employees) {
            LOGGER.info("test_read_xlsx_single_sheet :: Output - {}", emp);
        }
    }


    @Test
    public void test_read_xlsx_multiple_sheets() throws SpreadsheetReadException {
        // Excel Reader
        LOGGER.info("test_read_xlsx_multiple_sheets :: Reading file - {}", _filepath_xlsx_multiple_sheets);
        XlsxReader ger = new XlsxReader();

        // Read Sheet 1
        List<Employee> employees = ger.read(Employee.class, new File(_filepath_xlsx_multiple_sheets), 0);
        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);

        for (Employee emp : employees) {
            LOGGER.info("test_read_xlsx_multiple_sheets :: Output - {}", emp);
        }

        // Read Sheet 2
        List<Company> companies = ger.read(Company.class, new File(_filepath_xlsx_multiple_sheets), 1);
        Assert.assertNotNull(companies);
        Assert.assertTrue(companies.size() > 0);

        for (Company company : companies) {
            LOGGER.info("test_read_xlsx_multiple_sheets :: Output - {}", company);
        }
    }



    // Read from Stream

    @Test
    public void test_read_xlsx_single_sheet_from_stream() throws SpreadsheetReadException, FileNotFoundException {
        // Excel Reader
        LOGGER.info("test_read_xlsx_single_sheet_from_stream :: Reading file - {}", _filepath_xlsx_single_sheet);
        XlsxReader reader = new XlsxReader();

        // InputStream
        final InputStream fis = new FileInputStream(new File(_filepath_xlsx_single_sheet));

        // Read
        List<Employee> employees = reader.read(Employee.class, fis);
        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);

        for (Employee emp : employees) {
            LOGGER.info("test_read_xlsx_single_sheet :: Output - {}", emp);
        }
    }

    @Test
    public void test_read_xlsx_multiple_sheets_from_stream() throws SpreadsheetReadException, FileNotFoundException {
        // Excel Reader
        LOGGER.info("test_read_xlsx_multiple_sheets_from_stream :: Reading file - {}", _filepath_xlsx_multiple_sheets);
        XlsxReader reader = new XlsxReader();

        // InputStream
        final InputStream fisSheet1 = new FileInputStream(new File(_filepath_xlsx_multiple_sheets));

        // Read Sheet 1
        List<Employee> employees = reader.read(Employee.class, fisSheet1, 0);
        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);

        for (Employee emp : employees) {
            LOGGER.info("test_read_xlsx_multiple_sheets :: Output - {}", emp);
        }

        // InputStream
        final InputStream fisSheet2 = new FileInputStream(new File(_filepath_xlsx_multiple_sheets));

        // Read Sheet 2
        List<Company> companies = reader.read(Company.class, fisSheet2, 1);
        Assert.assertNotNull(companies);
        Assert.assertTrue(companies.size() > 0);

        for (Company company : companies) {
            LOGGER.info("test_read_xlsx_multiple_sheets :: Output - {}", company);
        }
    }


    // Read with Callback

    @Test
    public void test_read_xlsx_single_sheet_with_callback() throws SpreadsheetReadException {
        // Excel Reader
        LOGGER.info("test_read_xlsx_single_sheet_with_callback :: Reading file - {}", _filepath_xlsx_single_sheet);

        // file
        final File xlsxFile = new File(_filepath_xlsx_single_sheet);

        final List<Employee> employees = new ArrayList<Employee>();

        // Read
        XlsxReader reader = new XlsxReader();
        reader.read(Employee.class, xlsxFile, new RowListener<Employee>() {

            @Override
            public void row(int rowNum, Employee employee) {
                employees.add(employee);
                LOGGER.info("test_read_xlsx_single_sheet_with_callback :: Output - {}", employee);

            }
        });

        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);
    }



    // Read to Map


    @Test
    public void test_read_xlsx_as_Map() throws FileNotFoundException {
        // Excel Reader
        LOGGER.info("test_read_xlsx_as_Map :: Reading file - {}", _filepath_xlsx_single_sheet);
        XlsxReader ger = new XlsxReader();

        // Read
        /*
        List<Map<String, Object>> employees = ger.readAsMap(new File(_filepath_xlsx_single_sheet), 1);
        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);

        for (Map<String, Object> emp : employees) {
            LOGGER.info("test_read_xlsx_single_sheet :: Output - {}", emp);
        }
        */
    }


}
