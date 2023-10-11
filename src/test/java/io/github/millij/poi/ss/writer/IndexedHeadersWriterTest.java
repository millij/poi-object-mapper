package io.github.millij.poi.ss.writer;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import io.github.millij.bean.Employee;
import io.github.millij.bean.EmployeeIndexed;
import io.github.millij.poi.ss.model.annotations.SheetColumn;
import io.github.millij.poi.ss.writer.SpreadsheetWriter;
import io.github.millij.poi.ss.writer.SpreadsheetWriterTest;


public class IndexedHeadersWriterTest {
    private static final Logger LOGGER = LoggerFactory.getLogger(SpreadsheetWriterTest.class);

    private final String _path_test_output = "test-cases/output/";



    @Test
    public void test_write_xlsx_single_sheet() throws IOException {
        final String filepath_output_file = _path_test_output.concat("indexed_header_writer_sample.xlsx");

        // Excel Writer
        LOGGER.info("test_write_xlsx_single_sheet :: Writing to file - {}", filepath_output_file);
        SpreadsheetWriter gew = new SpreadsheetWriter(filepath_output_file);

        // Employees
        List<EmployeeIndexed> employees = new ArrayList<EmployeeIndexed>();
        employees.add(new EmployeeIndexed("1", "foo", 12, "MALE", 1.68, "Chennai"));
        employees.add(new EmployeeIndexed("2", "bar", 24, "MALE", 1.98, "Banglore"));
        employees.add(new EmployeeIndexed("3", "foo bar", 10, "FEMALE", 2.0, "Kolkata"));

        // Write
        gew.addSheet(EmployeeIndexed.class, employees);
        gew.write();
    }

}
