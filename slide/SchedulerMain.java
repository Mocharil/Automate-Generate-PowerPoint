package com.ebdesk.report.slide;

import java.util.Date;
import java.util.Timer;

//Main class
public class SchedulerMain {
	public static void main(String args[]) throws Exception {

		Timer time = new Timer(); // Instantiate Timer Object
		ScheduledTask st = new ScheduledTask(); // Instantiate SheduledTask class
		time.schedule(st, 0, 100000); // Create Repetitively task for every 1 secs
		Date now = new Date();
		//for demo only.
		for (int i = 0; i <= 5; i++) {
			System.out.println("Execution in Main Thread...." + i);
			Thread.sleep(101000);
			if (i == 5) {
				System.out.println("Application Terminates");
				System.exit(0);
			}
		}
	}
}