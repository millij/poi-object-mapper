package io.github.millij.poi.ss.reader;

import io.github.millij.bean.Emp_sameCols;
import io.github.millij.poi.SpreadsheetReadException;
import io.github.millij.poi.ss.reader.XlsReader;

import java.io.File;
import java.text.ParseException;
import java.util.List;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


public class ColsWithSameHeaderReadTest {

    private static final Logger LOGGER = LoggerFactory.getLogger(ColsWithSameHeaderReadTest.class);

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
        List<Emp_sameCols> employees = reader.read(Emp_sameCols.class, new File(_filepath_xls_single_sheet));
        Assert.assertNotNull(employees);
        Assert.assertTrue(employees.size() > 0);

        for (Emp_sameCols emp : employees) {
            LOGGER.info("test_read_xls_single_sheet :: Output - {}", emp);
        }
    }
}