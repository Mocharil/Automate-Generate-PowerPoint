package com.ebdesk.report.slide;

import java.awt.Color;
import java.io.*;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.TreeSet;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.axis.NumberAxis;
import org.jfree.chart.plot.XYPlot;
import org.jfree.chart.renderer.xy.XYItemRenderer;
import org.jfree.data.general.SeriesException;
import org.jfree.data.time.Day;
import org.jfree.data.time.Hour;
import org.jfree.data.time.Minute;
import org.jfree.data.time.Second;
import org.jfree.data.time.TimeSeries;
import org.jfree.data.time.TimeSeriesCollection;
import org.jfree.data.xy.XYDataset;

import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.DBObject;
import com.mongodb.Mongo;

import org.jfree.chart.ChartUtilities;

public class TimeSeriesChart {

	public void TimeSeries(final String Title) throws Exception {
		Mongo mongo = new Mongo("192.168.99.30:27017");
		DB db = mongo.getDB("social_media"); // file
		DBCollection collection = db.getCollection("twitter-post-search");

		ArrayList list = new ArrayList();
		DBCursor cursor = collection.find();
		while (cursor.hasNext()) {
			DBObject object2 = cursor.next();
			list.add(object2);
		}
		// ---------------time-----------------//
		ArrayList<String> settanggal = new ArrayList<String>(); // tanggal
		HashSet<String> settime1 = new HashSet<String>();
		Map<Integer, Integer> map = new HashMap<Integer, Integer>();

		DBObject c3 = (DBObject) list.get(0);
		String g = (String) c3.get("created_at");
		String g1;
		String[] time1;
		String[] time2 = g.split("\\s");
		String[] bulan = { "Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
		int nilai_bulan = 1;
		int tanggal1 = 1;
		int tanggal2 = 1;
		for (int bul = 0; bul < 12; bul++) {

			if (String.valueOf(time2[1]).equals(String.valueOf(bulan[bul]))) {
				nilai_bulan = (bul + 1) * 100;
				tanggal2 = nilai_bulan + Integer.parseInt(time2[2]) + Integer.parseInt(time2[5]) * 1000;
				map.put(tanggal2, 0);
				settanggal.add(String.valueOf(tanggal2));
				settime1.add(String.valueOf(tanggal2));
			}
		}

		int tanggal_awal = tanggal2;
		int index1 = 0;
		int index2 = 0;
		int tanggal_akhir = tanggal1;
		for (int ii = 0; ii < list.size(); ii++) {
			if (ii != 0) {
				DBObject c1 = (DBObject) list.get(ii);
				g1 = (String) c1.get("created_at");
				time1 = g1.split("\\s");
				for (int bul = 0; bul < 12; bul++) {

					if (String.valueOf(time1[1]).equals(String.valueOf(bulan[bul]))) {

						nilai_bulan = (bul + 1) * 100;
						tanggal1 = nilai_bulan + Integer.parseInt(time1[2]) + Integer.parseInt(time1[5]) * 1000;
						settime1.add(String.valueOf(tanggal1));

					}
				}
				settanggal.add(String.valueOf(tanggal1));
				if (tanggal_awal > tanggal1) {
					tanggal_awal = tanggal1;
					index1 = ii;
				}

				if (tanggal_akhir < tanggal1) {
					tanggal_akhir = tanggal1;
					index2 = ii;
				}

			}
			map.put(tanggal1, ii);
		}
		DBObject cawal = (DBObject) list.get(index1);
		DBObject cakhir = (DBObject) list.get(index2);
		String ga = (String) cawal.get("created_at");
		time1 = ga.split("\\s");
		String gk = (String) cakhir.get("created_at");
		time2 = gk.split("\\s");
		String[] time3;
		final TimeSeries series = new TimeSeries("Random Data");

		int indexmap;

		Object[] r = settime1.toArray();
		for (int oo = 0; oo < settime1.size(); oo++) {
			try {
				indexmap = map.get(Integer.valueOf((String) r[oo]));
				DBObject c1 = (DBObject) list.get(indexmap);
				g1 = (String) c1.get("created_at");
				time3 = g1.split("\\s");

				for (int bul = 0; bul < 12; bul++) {

					if (String.valueOf(time3[1]).equals(String.valueOf(bulan[bul]))) {

						nilai_bulan = (bul + 1);
					}
				}

				
				series.add(new Day(Integer.valueOf(time3[2]), nilai_bulan, Integer.valueOf(time3[5])),
						new Double(Collections.frequency(settanggal, String.valueOf(r[oo]))));
				
				

			} catch (Exception e) {
			}
		}

		final XYDataset dataset = (XYDataset) new TimeSeriesCollection(series);
		
		JFreeChart timechart = ChartFactory.createTimeSeriesChart("Trend Tweet for Keyword"+Title, "Time", "Tweet Count", dataset, false, false,
				false);
timechart.setBorderPaint(Color.blue);


		XYPlot plot = (XYPlot) timechart.getPlot();
		plot.setBackgroundPaint(Color.white);
		plot.setOutlinePaint(Color.CYAN);

		plot.setDomainGridlinePaint(Color.BLUE);
		

		int width = 580; /* Width of the image */
		int height = 370; /* Height of the image */
		File timeChart = new File("TimeChart.png");
		ChartUtilities.saveChartAsJPEG(timeChart, timechart, width, height);
		System.out.println("done");

	}
}
