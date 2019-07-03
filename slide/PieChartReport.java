package com.ebdesk.report.slide;

import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.GradientPaint;
import java.awt.Stroke;
import java.io.*;

import javax.swing.JFrame;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartPanel;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.block.BlockBorder;
import org.jfree.chart.labels.StandardPieSectionLabelGenerator;
import org.jfree.chart.plot.PieLabelLinkStyle;
import org.jfree.chart.plot.PiePlot;
import org.jfree.data.general.DefaultPieDataset;
import org.jfree.ui.RectangleEdge;
import org.jfree.ui.RectangleInsets;
import org.jfree.ui.VerticalAlignment;
import org.jfree.util.UnitType;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.plot.RingPlot;
import org.jfree.chart.title.LegendTitle;

public class PieChartReport {

	public void pie(String tweet_, float tweet, String reply_, float reply, String quoted_, float quoted,
			String retweet_, float retweet, String FILE, Color warna) throws IOException {

		DefaultPieDataset dpd = new DefaultPieDataset();
		dpd.setValue(tweet_, tweet);
		dpd.setValue(reply_, reply);
		dpd.setValue(quoted_, quoted);
		dpd.setValue(retweet_, retweet);

		JFreeChart freeChart = ChartFactory.createPieChart(null, dpd, true, false, false);
		// --------------background--------------//
		freeChart.setBackgroundPaint(new Color(255, 255, 255, 0));
		freeChart.setBorderPaint(Color.BLUE);

		PiePlot plot = (PiePlot) freeChart.getPlot();
		plot.setBackgroundPaint(new Color(255, 255, 255, 0));
		plot.setBaseSectionOutlinePaint(Color.WHITE);
		plot.setLabelGenerator(new StandardPieSectionLabelGenerator("{1}%"));
		plot.setSimpleLabelOffset(new RectangleInsets(UnitType.RELATIVE, 0.5, 0.5, 0.5, 0.5));
		plot.setSectionPaint(2, new Color(92, 177, 92));
		plot.setSectionPaint(1, new Color(59, 153, 181));
		plot.setSectionPaint(0, new Color(255, 85, 85));
		plot.setSectionPaint(3, new Color(255, 255, 93));
		plot.setLabelBackgroundPaint(null);
		plot.setLabelOutlinePaint(null);
		plot.setLabelShadowPaint(null);
		plot.setSectionOutlinesVisible(false);
		plot.setIgnoreZeroValues(true);
		plot.setLabelFont(new Font("Corbel", Font.BOLD, 20));

		plot.setOutlinePaint(new Color(255, 255, 255, 0));
		plot.setLabelBackgroundPaint(new Color(255, 255, 255, 0));
		plot.setLabelOutlinePaint(new Color(255, 255, 255, 0));
		plot.setLabelPaint(Color.BLACK);
		plot.setShadowPaint(new Color(255, 255, 255, 0));
		plot.setOutlineVisible(false);
		plot.setLabelLinkStyle(PieLabelLinkStyle.QUAD_CURVE);
		plot.setIgnoreZeroValues(true);

		LegendTitle legend = freeChart.getLegend();
		legend.setFrame(BlockBorder.NONE);
		legend.setPosition(RectangleEdge.LEFT);
		legend.setVerticalAlignment(VerticalAlignment.BOTTOM);
		legend.setBackgroundPaint(new Color(255, 255, 255, 0));
		legend.setItemFont(new Font("Corbel", Font.BOLD, 18));

		int width = 640;
		int height = 480;
		File pieChart = new File(FILE);
		ChartUtilities.saveChartAsPNG(pieChart, freeChart, width, height);
		System.out.println("selesai");
	}

	public void ringchart(String tweet_, float tweet, String reply_, float reply, String quoted_, float quoted,
			String retweet_, float retweet, String FILE, Color warna) throws IOException {
		JFrame f = new JFrame("Chart");

		DefaultPieDataset dataset = new DefaultPieDataset();
		dataset.setValue(tweet_, tweet);
		dataset.setValue(reply_, reply);
		dataset.setValue(quoted_, quoted);
		dataset.setValue(retweet_, retweet);

		JFreeChart chart = ChartFactory.createRingChart(null, dataset, true, false, false);
		chart.setBackgroundPaint(new Color(255, 255, 255, 0));
		chart.setBorderPaint(new Color(255, 255, 255, 0));

		RingPlot pie = (RingPlot) chart.getPlot();

		pie.setBackgroundPaint(new Color(255, 255, 255, 0));
		pie.setOutlineVisible(false);
		pie.setShadowPaint(null);
		pie.setSimpleLabels(true);
		pie.setLabelGenerator(new StandardPieSectionLabelGenerator("{1}%"));
		pie.setSimpleLabelOffset(new RectangleInsets(UnitType.RELATIVE, 0.5, 0.5, 0.5, 0.5));
		pie.setSectionPaint(1, new Color(255, 85, 85));
		pie.setSectionPaint(0, new Color(59, 153, 181));
		pie.setSectionPaint(2, new Color(92, 177, 92));
		pie.setSectionPaint(3, new Color(255, 255, 93));
		pie.setLabelBackgroundPaint(null);
		pie.setLabelOutlinePaint(null);
		pie.setLabelShadowPaint(null);
		pie.setSectionDepth(0.33);
		pie.setSectionOutlinesVisible(false);
		pie.setSeparatorsVisible(true);
		pie.setSeparatorPaint(warna);
		pie.setSeparatorStroke(new BasicStroke(4));
		pie.setIgnoreZeroValues(true);
		pie.setLabelFont(new Font("Corbel", Font.BOLD, 20));

		LegendTitle legend = chart.getLegend();
		legend.setFrame(BlockBorder.NONE);
		legend.setPosition(RectangleEdge.LEFT);
		legend.setVerticalAlignment(VerticalAlignment.BOTTOM);
		legend.setBackgroundPaint(new Color(255, 255, 255, 0));
		legend.setItemFont(new Font("Corbel", Font.BOLD, 18));

		f.add(new ChartPanel(chart) {
			public Dimension getPreferredSize() {
				return new Dimension(400, 400);
			}
		});

		File pieChart = new File(FILE);
		ChartUtilities.saveChartAsPNG(pieChart, chart, 640, 480);
		System.out.println("selesai");

	}

}