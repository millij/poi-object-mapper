package io.github.millij.dates;

import java.io.FileNotFoundException;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;

import io.github.millij.bean.Schedule;
//import io.github.millij.bean.Schedules;
import io.github.millij.poi.ss.writer.SpreadsheetWriter;

//import io.github.millij.bean.Schedules;

public class DateTest {

	@Test
	public void writeDatesTest() throws IOException {

		List<Schedule> schedules = new ArrayList<Schedule>();
		final String filepath_output_file = "src/test/resources/sample-files/write_formatted_date_sample.xls";
		SpreadsheetWriter gew = new SpreadsheetWriter(filepath_output_file);

		// Schedules
		schedules.add(new Schedule("Friday", "02/11/2001"));
		schedules.add(new Schedule("Saturday", "16/09/2023"));

		gew.addSheet(Schedule.class, schedules);
		gew.write();
	}

}
