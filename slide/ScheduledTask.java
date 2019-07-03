package com.ebdesk.report.slide;

import java.util.TimerTask;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;
import java.util.TimeZone;

public class ScheduledTask extends TimerTask {

	Date now; // to display current time

	// Add your task here
	public void run() {
		now = new Date(); // initialize date
		System.out.println("Time is :" + now); // Display current time

		SimpleDateFormat dateFormat = new SimpleDateFormat("dd MMMM yyyy mm", Locale.ENGLISH);
		dateFormat.setTimeZone(TimeZone.getTimeZone("UTC"));

		SNATwitterAccount report = new SNATwitterAccount("slideshowTwitterAccount" + dateFormat.format(now) + ".ppt", 2,
				960, 540, "ShamsiAli2", "10", "ShamsiAli2");
		try {
			report.execute();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
}