package com.github.millij.eom;

import java.text.ParseException;
import java.util.List;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.millij.eom.bean.Company;
import com.github.millij.eom.bean.Employee;


public class GenericExcelReaderTest {

    private static final Logger LOGGER = LoggerFactory.getLogger(GenericExcelReaderTest.class);

    // XLS
    private String _filepath_xls_single_sheet;
    private String _filepath_xls_multiple_sheets;

    // XLSX
    private String _filepath_xlsx_single_sheet;
    private String _filepath_xlsx_multiple_sheets;

    // Setup
    // ------------------------------------------------------------------------

    @Before
    public void setup() throws ParseException {
        // filepaths

        // xls
        _filepath_xls_single_sheet = "src/test/resources/sample-files/xls_sample_single_sheet.xls";
        _filepath_xls_multiple_sheets = "src/test/resources/sample-files/xls_sample_multiple_sheets.xls";

        // xlsx
        _filepath_xlsx_single_sheet = "src/test/resources/sample-files/xlsx_sample_single_sheet.xlsx";
        _filepath_xlsx_multiple_sheets = "src/test/resources/sample-files/xlsx_sample_multiple_sheets.xlsx";
    }

    @After
    public void teardown() {
        // nothing to do
    }


    // Tests
    // ------------------------------------------------------------------------

    @Test
    public void test_read_xlsx_single_sheet() {
        // Excel Reader
        LOGGER.info("test_read_xlsx_single_sheet :: Reading file - {}", _filepath_xlsx_single_sheet);
        GenericExcelReader ger = new GenericExcelReader(_filepath_xlsx_single_sheet);

        // Read
        List<Employee> employees = ger.read(Employee.class);
        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);

        for (Employee emp : employees) {
            LOGGER.info("test_read_xlsx_single_sheet :: Output - {}", emp);
        }
    }


    @Test
    public void test_read_xlsx_multiple_sheets() {
        // Excel Reader
        LOGGER.info("test_read_xlsx_multiple_sheets :: Reading file - {}", _filepath_xlsx_multiple_sheets);
        GenericExcelReader ger = new GenericExcelReader(_filepath_xlsx_multiple_sheets);

        // Read Sheet 1
        List<Employee> employees = ger.read(0, Employee.class);
        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);

        for (Employee emp : employees) {
            LOGGER.info("test_read_xlsx_multiple_sheets :: Output - {}", emp);
        }

        // Read Sheet 2
        List<Company> companies = ger.read(1, Company.class);
        Assert.assertNotNull(companies);
        Assert.assertTrue(companies.size() > 0);

        for (Company company : companies) {
            LOGGER.info("test_read_xlsx_multiple_sheets :: Output - {}", company);
        }
    }

}
