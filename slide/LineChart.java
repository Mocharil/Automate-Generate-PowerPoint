package com.ebdesk.report.slide;

import org.jfree.chart.ChartPanel;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartFrame;
import org.jfree.chart.JFreeChart;
import org.jfree.ui.ApplicationFrame;
import org.jfree.ui.RefineryUtilities;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.PiePlot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.data.category.DefaultCategoryDataset;

import java.awt.Color;
import java.io.*;
import org.jfree.chart.ChartUtilities;

public class LineChart {
	public void Line(String title) throws IOException {
		DefaultCategoryDataset line_chart_dataset = new DefaultCategoryDataset();
		for (int hitungan = 0; hitungan <= 1000; hitungan++) {
			
		
		line_chart_dataset.addValue(hitungan+3, "trend", Integer.toString(hitungan));}
		
//		line_chart_dataset.addValue(30, "trend", 2);
//		line_chart_dataset.addValue(60, "trend", 4);
//		line_chart_dataset.addValue(120, "trend", 5);
//		line_chart_dataset.addValue(240, "trend", 6);
//		line_chart_dataset.addValue(300, "trend", 7);
		JFreeChart lineChartObject = ChartFactory.createLineChart("Trend Tweet for Keyword: " + title, "Time",
				"Tweet Count", line_chart_dataset, PlotOrientation.VERTICAL, false, false, false);

		ChartPanel chartPanel = new ChartPanel(lineChartObject);
		lineChartObject.setBackgroundPaint(Color.white);
		CategoryPlot plot = lineChartObject.getCategoryPlot();
		plot.setBackgroundPaint(Color.white);
		plot.setDomainGridlinePaint(Color.gray);
		plot.setRangeGridlinePaint(Color.gray);
//
//		int width = 640; /* Width of the image */
//		int height = 480; /* Height of the image */
//		File lineChart = new File("Trend.png");
//		ChartUtilities.saveChartAsJPEG(lineChart, lineChartObject, width, height);
		//
		 ChartFrame cf = new ChartFrame("Tweets by Type", lineChartObject);
		 cf.setSize(1000, 1000);
		 cf.setVisible(true);
		 cf.setLocationRelativeTo(null);
	}
}
