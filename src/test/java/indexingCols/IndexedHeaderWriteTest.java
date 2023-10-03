package indexingCols;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import io.github.millij.bean.Emp_indexed;
import io.github.millij.bean.Employee;
import io.github.millij.poi.ss.model.annotations.SheetColumn;
import io.github.millij.poi.ss.writer.SpreadsheetWriter;
import io.github.millij.poi.ss.writer.SpreadsheetWriterTest;

public class IndexedHeaderWriteTest
{
	private static final Logger LOGGER = LoggerFactory.getLogger(SpreadsheetWriterTest.class);

    private final String _path_test_output = "test-cases/output/";
	
	
	
    @Test
    public void test_write_xlsx_single_sheet() throws IOException {
        final String filepath_output_file = _path_test_output.concat("indexed_header_sample.xlsx");

        // Excel Writer
        LOGGER.info("test_write_xlsx_single_sheet :: Writing to file - {}", filepath_output_file);
        SpreadsheetWriter gew = new SpreadsheetWriter(filepath_output_file);

        // Employees
        List<Emp_indexed> employees = new ArrayList<Emp_indexed>();
        employees.add(new Emp_indexed("1", "foo", 12, "MALE", 1.68,"Chennai"));
        employees.add(new Emp_indexed("2", "bar", 24, "MALE", 1.98,"Banglore"));
        employees.add(new Emp_indexed("3", "foo bar", 10, "FEMALE",2.0,"Kolkata"));

        // Write
        gew.addSheet(Emp_indexed.class, employees);
        gew.write();
    }

}
