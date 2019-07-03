package com.ebdesk.report.slide;

import java.awt.Color;
import java.awt.GradientPaint;
import java.awt.Graphics2D;
import java.awt.SystemColor;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;

import javax.swing.JTable;
import javax.swing.table.JTableHeader;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.plot.Plot;
import org.jfree.chart.plot.PlotOrientation;
import org.jfree.chart.plot.XYPlot;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.renderer.category.StandardBarPainter;
import org.jfree.data.category.DefaultCategoryDataset;
import org.jfree.chart.ChartUtilities;

public class BarChart {

	public ArrayList<String> bar(ArrayList<String> listhtg, HashSet<String> sethtg, String name_file, String Caption,
			int a, int choice, String warna) throws Exception {

		Map<Integer, Integer> maphtg = new HashMap<Integer, Integer>();
		ArrayList<Integer> listnilaihtg = new ArrayList<Integer>();
		ArrayList<String> ranking = new ArrayList<String>();
		Object[] rz = sethtg.toArray();
		for (int x = 0; x < sethtg.size(); x++) {
			maphtg.put(x, Collections.frequency(listhtg, String.valueOf(rz[x])));
			listnilaihtg.add(Collections.frequency(listhtg, String.valueOf(rz[x])));
		}

		final DefaultCategoryDataset datasethtg = new DefaultCategoryDataset();
		Collections.sort(listnilaihtg, Collections.reverseOrder());
		for (int xx = 0; xx < a; xx++) {
			for (int xy = 0; xy < sethtg.size(); xy++) {
				if (listnilaihtg.get(xx).equals(maphtg.get(xy))) {
					datasethtg.addValue(listnilaihtg.get(xx), "Nama", String.valueOf(rz[xy]));
					ranking.add(String.valueOf(rz[xy]));
				}
			}
		}

		if (choice == 1) {
			if (a < ranking.size()) {
				for (int g = 1; g < ranking.size() - a; g++) {
					try {
						datasethtg.removeColumn(a);
					} catch (Exception e) {
					}

				}
			}
			File BarChart2 = new File(name_file);
			JFreeChart barChart2 = ChartFactory.createBarChart(Caption, null, null, datasethtg,
					PlotOrientation.HORIZONTAL, false, false, false);
			save(barChart2, BarChart2, warna);
		}

		else if (choice == 2) {
			if (a < ranking.size()) {
				for (int g = 1; g < ranking.size() - a; g++) {
					try {
						datasethtg.removeColumn(a);
					} catch (Exception e) {
					}

				}
			}
			File BarChart2 = new File(name_file);
			JFreeChart barChart2 = ChartFactory.createBarChart(Caption, null, null, datasethtg,
					PlotOrientation.VERTICAL, false, false, false);
			save(barChart2, BarChart2, warna);

		}

		System.out.println("DONE bar " + Caption);
		return (ranking);

	}

	private static void save(JFreeChart barChart2, File BarChart2, String warna) throws IOException {
		int width = 640; /* Width of the image */
		int height = 480; /* Height of the image */
		barChart2.setBackgroundPaint(new Color(255, 255, 255, 0));

		CategoryPlot cplot = (CategoryPlot) barChart2.getPlot();
		cplot.setBackgroundPaint(new Color(255, 255, 255, 0));
		cplot.setOutlinePaint(new Color(255, 255, 255, 0));
		// set bar chart color

		((BarRenderer) cplot.getRenderer()).setBarPainter(new StandardBarPainter());

		BarRenderer r = (BarRenderer) barChart2.getCategoryPlot().getRenderer();
		if (warna.equals("orange")) {

			r.setSeriesPaint(0, new Color(243, 108, 61));
		} else if (warna.equals("red")) {
			r.setSeriesPaint(0, new Color(255, 0, 0));
		} else if (warna.equals("green")) {
			r.setSeriesPaint(0, new Color(112, 173, 71));
		} else {
			r.setSeriesPaint(0, new Color(61, 140, 255));
		}

		ChartUtilities.saveChartAsPNG(BarChart2, barChart2, width, height);

	}
}