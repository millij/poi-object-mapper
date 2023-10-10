package io.github.millij.dates;

import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.junit.Test;

import io.github.millij.bean.Schedule;
// import io.github.millij.bean.Schedules;
import io.github.millij.poi.ss.writer.SpreadsheetWriter;

// import io.github.millij.bean.Schedules;


public class DateTest {

    @Test
    public void writeDatesTest() throws IOException, ParseException {

        List<Schedule> schedules = new ArrayList<Schedule>();
        final String filepath_output_file = "src/test/resources/sample-files/write_formatted_date_sample.xls";
        SpreadsheetWriter gew = new SpreadsheetWriter(filepath_output_file);

        SimpleDateFormat formatter = new SimpleDateFormat("dd/MM/yyyy");
        DateTimeFormatter format = DateTimeFormatter.ofPattern("dd/MM/yyyy");

        Date date1 = formatter.parse("02/05/2001");
        Date date2 = formatter.parse("03/06/2001");
        LocalDate ldate1 = LocalDate.parse("02/05/2002", format);
        LocalDate ldate2 = LocalDate.parse("03/06/2002", format);

        // Schedules
        schedules.add(new Schedule("Friday", date1, ldate1));
        schedules.add(new Schedule("Saturday", date2, ldate2));

        gew.addSheet(Schedule.class, schedules);
        gew.write();
    }

}
