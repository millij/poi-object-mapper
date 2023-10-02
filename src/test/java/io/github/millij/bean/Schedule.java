package io.github.millij.bean;

//import java.util.Date;

import io.github.millij.poi.ss.model.annotations.Sheet;
import io.github.millij.poi.ss.model.annotations.SheetColumn;

@Sheet("Schedules")
public class Schedule {

	@SheetColumn("Day")
	private String day;

	@SheetColumn(value = "Date", isFormatted = true, format = "yyyy/MM/dd")
	private String date;

	public Schedule(String day, String date) {
		super();
		this.day = day;
		this.date = date;
	}

	public String getDay() {
		return day;
	}

	public void setDay(String day) {
		this.day = day;
	}

	public String getDate() {
		return date;
	}

	public void setDate(String date) {
		this.date = date;
	}

	@Override
	public String toString() {
		return "Schedules [day=" + day + ", date=" + date + "]";
	}

}
