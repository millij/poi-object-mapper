package com.github.millij.poi.ss.reader;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.ParseException;
import java.util.List;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.millij.bean.Company;
import com.github.millij.bean.Employee;
import com.github.millij.poi.ExcelReadException;
import com.github.millij.poi.ss.reader.XlsxReader;


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
    

    // XLSX

    @Test
    public void test_read_xlsx_single_sheet() throws IOException, ExcelReadException {
        // Excel Reader
        LOGGER.info("test_read_xlsx_single_sheet :: Reading file - {}", _filepath_xlsx_single_sheet);
        XlsxReader reader = new XlsxReader();

        // Read
        List<Employee> employees = reader.read(new File(_filepath_xlsx_single_sheet), Employee.class);
        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);

        for (Employee emp : employees) {
            LOGGER.info("test_read_xlsx_single_sheet :: Output - {}", emp);
        }
    }


    @Test
    public void test_read_xlsx_multiple_sheets() throws IOException, ExcelReadException {
        // Excel Reader
        LOGGER.info("test_read_xlsx_multiple_sheets :: Reading file - {}", _filepath_xlsx_multiple_sheets);
        XlsxReader ger = new XlsxReader();

        // Read Sheet 1
        List<Employee> employees = ger.read(new File(_filepath_xlsx_multiple_sheets), 0, Employee.class);
        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);

        for (Employee emp : employees) {
            LOGGER.info("test_read_xlsx_multiple_sheets :: Output - {}", emp);
        }

        // Read Sheet 2
        List<Company> companies = ger.read(new File(_filepath_xlsx_multiple_sheets), 1, Company.class);
        Assert.assertNotNull(companies);
        Assert.assertTrue(companies.size() > 0);

        for (Company company : companies) {
            LOGGER.info("test_read_xlsx_multiple_sheets :: Output - {}", company);
        }
    }



    // Read to Map

    @Test
    public void test_read_xlsx_as_Map() throws ExcelReadException, FileNotFoundException {
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
