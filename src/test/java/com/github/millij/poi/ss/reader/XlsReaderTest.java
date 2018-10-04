package com.github.millij.poi.ss.reader;

import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.util.List;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Ignore;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.millij.bean.Company;
import com.github.millij.bean.Employee;
import com.github.millij.poi.ExcelReadException;


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
    

    // XLS

    @Ignore
    @Test
    public void test_read_xls_single_sheet() throws ExcelReadException, IOException {
        // Excel Reader
        LOGGER.info("test_read_xls_single_sheet :: Reading file - {}", _filepath_xls_single_sheet);
        XlsReader ger = new XlsReader();

        // Read
        List<Employee> employees = ger.read(new File(_filepath_xls_single_sheet), Employee.class);
        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);

        for (Employee emp : employees) {
            LOGGER.info("test_read_xls_single_sheet :: Output - {}", emp);
        }
    }

    @Ignore
    @Test
    public void test_read_xls_multiple_sheets() throws ExcelReadException {
        // Excel Reader
        LOGGER.info("test_read_xlsx_multiple_sheets :: Reading file - {}", _filepath_xls_multiple_sheets);
        XlsReader ger = new XlsReader();

        // Read Sheet 1
        List<Employee> employees = ger.read(new File(_filepath_xls_multiple_sheets), 0, Employee.class);
        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);

        for (Employee emp : employees) {
            LOGGER.info("test_read_xls_multiple_sheets :: Output - {}", emp);
        }

        // Read Sheet 2
        List<Company> companies = ger.read(new File(_filepath_xls_multiple_sheets), 1, Company.class);
        Assert.assertNotNull(companies);
        Assert.assertTrue(companies.size() > 0);

        for (Company company : companies) {
            LOGGER.info("test_read_xls_multiple_sheets :: Output - {}", company);
        }
    }


}
