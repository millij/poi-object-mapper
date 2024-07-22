package io.github.millij.poi.ss.reader;

import java.io.File;
import java.text.ParseException;
import java.util.List;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import io.github.millij.bean.DataTypesBean;
import io.github.millij.poi.SpreadsheetReadException;


public class XlsxReaderDataTypesTest {

    private static final Logger LOGGER = LoggerFactory.getLogger(XlsxReaderDataTypesTest.class);

    // XLSX
    private String _filepath_xlsx_data_types;


    // Setup
    // ------------------------------------------------------------------------

    @Before
    public void setup() throws ParseException {
        // sample files
        _filepath_xlsx_data_types = "src/test/resources/sample-files/xlsx_sample_data_types.xlsx";
    }

    @After
    public void teardown() {
        // nothing to do
    }


    // Tests
    // ------------------------------------------------------------------------


    // Read from file

    @Test
    public void test_read_xlsx_data_types() throws SpreadsheetReadException {
        // Excel Reader
        LOGGER.info("test_read_xlsx_data_types :: Reading file - {}", _filepath_xlsx_data_types);
        XlsxReader reader = new XlsxReader();

        // Read
        List<DataTypesBean> beans = reader.read(DataTypesBean.class, new File(_filepath_xlsx_data_types));
        Assert.assertNotNull(beans);
        Assert.assertTrue(beans.size() > 0);

        for (final DataTypesBean bean : beans) {
            LOGGER.info("test_read_xlsx_data_types :: Output - {}", bean);
        }
    }


}
