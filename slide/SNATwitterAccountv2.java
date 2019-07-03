package com.ebdesk.report.slide;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.sql.Date;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;

import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.TimeZone;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.imageio.ImageIO;

import org.apache.poi.hslf.usermodel.HSLFAutoShape;
import org.apache.poi.hslf.usermodel.HSLFFill;
import org.apache.poi.hslf.usermodel.HSLFPictureData;
import org.apache.poi.hslf.usermodel.HSLFPictureShape;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTextBox;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.bson.Document;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.XYPlot;
import org.jfree.data.time.Day;
import org.jfree.data.time.TimeSeries;
import org.jfree.data.time.TimeSeriesCollection;
import org.jfree.data.xy.XYDataset;

import com.kennycason.kumo.WordFrequency;
import com.mongodb.MongoClient;
import com.mongodb.MongoCredential;
import com.mongodb.ServerAddress;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoCursor;
import com.mongodb.client.MongoDatabase;

import javafx.application.Application;
import javafx.embed.swing.SwingFXUtils;
import javafx.scene.Scene;
import javafx.scene.SnapshotParameters;
import javafx.scene.chart.PieChart;
import javafx.scene.image.WritableImage;
import javafx.scene.layout.StackPane;
import javafx.stage.Stage;

public class SNATwitterAccountv2 {

	private int version;
	private String filename;
	private int width;
	private int height;
	private String tema;

	public SNATwitterAccountv2(String filename, int version, int width, int height, String tema) {
		this.filename = filename;
		this.version = version;
		this.height = height;
		this.width = width;
		this.tema = tema;
	}

	public void execute() throws Exception {

		HSLFSlideShow ppt = new HSLFSlideShow();
		ppt.setPageSize(new java.awt.Dimension(width, height));

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\twitter_logo.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(920, 0, 40, 40));

		// Slide 1
		HSLFSlide slide1 = ppt.createSlide();
		slideMaster(slide1);

		// Slide 2
		HSLFSlide slide2 = ppt.createSlide();
		slideInformation(ppt, slide2);
		slide2.addShape(pictNew);

		// Slide 3
		HSLFSlide slide3 = ppt.createSlide();
		slideTrendTweet(ppt, slide3);
		slide3.addShape(pictNew);

		// Slide 4
		HSLFSlide slide31 = ppt.createSlide();
		slideTrendTweetperHour(ppt, slide31);
		slide31.addShape(pictNew);

		// Slide 4
		HSLFSlide slide4 = ppt.createSlide();
		slideOptimalTimeDiagram(ppt, slide4);
		slide4.addShape(pictNew);

		// Slide 5
		HSLFSlide slide5 = ppt.createSlide();
		slideTopikPerbicaraan(slide5);
		slide5.addShape(pictNew);

		// Slide 6
		HSLFSlide slide6 = ppt.createSlide();
		slideTweetsType(ppt, slide6);
		slide6.addShape(pictNew);

		// Slide 7
		HSLFSlide slide7 = ppt.createSlide();
		slideTweetsResponse(ppt, slide7);
		slide7.addShape(pictNew);

		// Slide 8
		HSLFSlide slide8 = ppt.createSlide();
		slideTopEngagedAccounts(slide8);
		slide8.addShape(pictNew);

		// Slide 9
		HSLFSlide slide9 = ppt.createSlide();
		slideTopSharedLinks(slide9);
		slide9.addShape(pictNew);

		// Slide 10
		HSLFSlide slide10 = ppt.createSlide();
		slideTopEngagedAccount2(slide10);
		slide10.addShape(pictNew);

		// Slide 12
		HSLFSlide slide12 = ppt.createSlide();
		slideTopTweets(slide12);
		slide12.addShape(pictNew);

		// Slide 13
		HSLFSlide slide13 = ppt.createSlide();
		slideNetworkofRetweet(slide13);
		slide13.addShape(pictNew);

		// Slide 8
		HSLFSlide slide22 = ppt.createSlide();
		slideConclusion(slide22);

		FileOutputStream out = new FileOutputStream(filename);
		ppt.write(out);
		out.close();
		System.out.println("-------------------> Done <-------------------");
	}

	private void setBackground(HSLFSlide slide, int ver) {
		if (ver == 1) {
			HSLFShape shape = new HSLFAutoShape(ShapeType.RECT);
			shape.setAnchor(new java.awt.Rectangle(0, 0, width, height));
			HSLFFill fill = shape.getFill();
			fill.setBackgroundColor(Color.BLACK);
			fill.setForegroundColor(Color.BLACK);
			slide.addShape(shape);
		}

	}

	private void setTextColor(HSLFTextRun text, int ver) {
		if (ver == 1) {
			text.setFontColor(Color.WHITE);
		} else {
			text.setFontColor(Color.BLACK);
		}
	}

	public void judul(HSLFSlide slide, String atas, int versi) {

		HSLFTextBox title = new HSLFTextBox();
		title.setText(atas);
		title.setAnchor(new java.awt.Rectangle(50, 20, 860, 100));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(36.);
		runTitle.setFontFamily("Corbel (Headings)");

		if (versi == 1) {
			runTitle.setFontColor(Color.BLACK);
		} else {
			runTitle.setFontColor(Color.WHITE);
		}

		slide.addShape(title);

		// ========garis==========//
		HSLFTextBox contentx1 = new HSLFTextBox();
		contentx1.setAnchor(new java.awt.Rectangle(0, 50, 250, 440));
		contentx1.setFillColor(new Color(169, 165, 124));
		slide.addShape(contentx1);

		HSLFTextBox contentx12 = new HSLFTextBox();
		contentx12.setAnchor(new java.awt.Rectangle(920, 50, 40, 440));
		contentx12.setFillColor(new Color(228, 228, 228));
		slide.addShape(contentx12);

	}

	public void slideMaster(HSLFSlide slide) {
		setBackground(slide, version);
		// =====================garis====================//
		HSLFTextBox contentx = new HSLFTextBox();
		contentx.setAnchor(new java.awt.Rectangle(755, 50, 205, 440));
		contentx.setRotation(180);
		contentx.setFillColor(new Color(227, 227, 227));
		slide.addShape(contentx);

		HSLFTextBox contentx1 = new HSLFTextBox();
		contentx1.setAnchor(new java.awt.Rectangle(0, 50, 750, 440));
		contentx1.setFillColor(new Color(169, 165, 124));
		slide.addShape(contentx1);

		// ===========judul============//
		HSLFTextBox title = new HSLFTextBox();
		title.setText("SOCIAL");
		title.appendText("NETWORK", true);
		title.appendText("ANALYSIS", true);
		title.setAnchor(new java.awt.Rectangle(60, 150, 950, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.LEFT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);
		titleP.setLineSpacing(85.0);
		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(59.);
		runTitle.setFontFamily("Corbel (Headings)");

		runTitle.setFontColor(Color.red);

		HSLFTextParagraph titleP1 = title.getTextParagraphs().get(2);
		HSLFTextRun runTitle1 = titleP1.getTextRuns().get(0);
		runTitle1.setFontSize(59.);
		runTitle1.setFontFamily("Corbel (Headings)");

		runTitle1.setFontColor(Color.WHITE);
		slide.addShape(title);

		// --------------------Investigation--------------------
		HSLFTextBox issue = new HSLFTextBox();
		issue.setText("Twitter - @" + tema);
		issue.setAnchor(new java.awt.Rectangle(60, 360, 900, 50));
		HSLFTextParagraph issueP = issue.getTextParagraphs().get(0);
		issueP.setAlignment(TextAlign.LEFT);
		HSLFTextRun runIssue = issueP.getTextRuns().get(0);
		runIssue.setFontSize(22.);
		runIssue.setFontFamily("Corbel (Body)");
		runIssue.setItalic(true);
		runIssue.setFontColor(Color.WHITE);
		slide.addShape(issue);

	}

	public void slideInformation(HSLFSlideShow ppt, HSLFSlide slide) throws IOException {
		setBackground(slide, version);
		// ========garis==========//
		HSLFTextBox contentx1 = new HSLFTextBox();
		contentx1.setAnchor(new java.awt.Rectangle(0, 50, 250, 440));
		contentx1.setFillColor(new Color(169, 165, 124));
		slide.addShape(contentx1);

		HSLFTextBox contentx12 = new HSLFTextBox();
		contentx12.setAnchor(new java.awt.Rectangle(920, 50, 40, 440));
		contentx12.setFillColor(new Color(228, 228, 228));
		slide.addShape(contentx12);

		// ==========judul===========//
		HSLFTextBox content = new HSLFTextBox();
		content.setText("Information");
		content.setAnchor(new java.awt.Rectangle(20, 60, 630, 420));
		HSLFTextParagraph contentP2x = content.getTextParagraphs().get(0);
		contentP2x.setAlignment(TextAlign.LEFT);

		HSLFTextRun runContent2x = contentP2x.getTextRuns().get(0);
		runContent2x.setFontSize(54.);
		runContent2x.setFontFamily("Haettenschweiler");
		runContent2x.setFontColor(Color.BLACK);
		slide.addShape(content);

		// ------------- add picture -----------//
		HSLFPictureData pd3 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\photoTWakun.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3 = new HSLFPictureShape(pd3);
		pictNew3.setAnchor(new java.awt.Rectangle(0, 140, 250, 250));
		slide.addShape(pictNew3);

		// -------------------------Content2------------------------//

		HSLFTextBox content2 = new HSLFTextBox();
		content2.setText("Full Name\t\t\t:");
		content2.appendText("\nUsername\t\t\t: ", false);
		content2.appendText("\nBiography\t\t\t: ", false);
		content2.appendText("\nUrl\t\t\t: ", false);
		content2.appendText("\nLocation\t\t\t: ", false);
		content2.appendText("\nJoined Date\t\t: ", false);
		content2.appendText("\nBirth Date\t\t\t: ", false);
		content2.appendText("\nFollowings Count\t\t: ", false);
		content2.appendText("\nFollowers Count\t\t: ", false);
		content2.appendText("\nTweets Count\t\t: ", false);
		content2.appendText("\nLikes Count\t\t: ", false);
		content2.appendText("\nLists Count\t\t: ", false);
		content2.setAnchor(new java.awt.Rectangle(275, 50, 630, 260));
		content2.setLineColor(new Color(191, 191, 191));
		content2.setLineWidth(2);

		HSLFTextParagraph contentP2 = content2.getTextParagraphs().get(0);
		contentP2.setAlignment(TextAlign.LEFT);
		contentP2.setSpaceAfter(0.);
		contentP2.setSpaceBefore(0.);
		contentP2.setLineSpacing(110.0);

		HSLFTextRun runContent2 = contentP2.getTextRuns().get(0);
		runContent2.setFontSize(16.);
		runContent2.setFontFamily("Corbel (Body)");
		runContent2.setFontColor(Color.BLACK);
		slide.addShape(content2);
	}

	public void slideTrendTweet(HSLFSlideShow ppt, HSLFSlide slide) throws IOException {
		setBackground(slide, version);
		judul(slide, "Trend Tweet", 1);

		// ------------- add picture -----------//
		HSLFPictureData pd3a = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\TimeChartTWAccount.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3a = new HSLFPictureShape(pd3a);
		pictNew3a.setAnchor(new java.awt.Rectangle(270, 90, 610, 300));
		slide.addShape(pictNew3a);

		// ============content2===========//
		HSLFTextBox content12 = new HSLFTextBox();
		content12.setText("Total Tweet : " + " tweets");
		content12.appendText("\nTimeframe : " + " days", false);

		content12.setAnchor(new java.awt.Rectangle(270, 430, 240, 55));
		content12.setLineColor(new Color(191, 200, 196));
		content12.setLineWidth(2);

		HSLFTextParagraph contentP12 = content12.getTextParagraphs().get(0);
		contentP12.setAlignment(TextAlign.JUSTIFY);
		contentP12.setSpaceAfter(2.);
		contentP12.setSpaceBefore(10.);
		contentP12.setLineSpacing(110.0);

		HSLFTextRun runContent12 = contentP12.getTextRuns().get(0);
		runContent12.setFontSize(15.);
		runContent12.setFontFamily("Calibri");
		runContent12.setBold(true);
		setTextColor(runContent12, version);
		slide.addShape(content12);

		// ============content3===========//
		HSLFTextBox content123w = new HSLFTextBox();
		content123w.setText("Max tweet= " + "tweets/day");
		content123w.appendText("\nMin tweet= " + "tweets/day", false);
		content123w.appendText("\nAv. tweet= " + "tweets/day", false);
		content123w.appendText("\nDalam sehari, " + tema + " biasanya mengirimkan 1 tweet.", false);
		content123w.setAnchor(new java.awt.Rectangle(5, 70, 235, 140));
		content123w.setLineColor(new Color(150, 150, 150));
		content123w.setLineWidth(2);
		content123w.setFillColor(Color.WHITE);

		HSLFTextParagraph contentP123w = content123w.getTextParagraphs().get(0);
		contentP123w.setAlignment(TextAlign.JUSTIFY);
		contentP123w.setSpaceAfter(2.);
		contentP123w.setSpaceBefore(10.);
		contentP123w.setLineSpacing(110.0);

		HSLFTextRun runContent123w = contentP123w.getTextRuns().get(0);
		runContent123w.setFontSize(16.);
		runContent123w.setFontFamily("Corbel (Body)");
		setTextColor(runContent123w, version);
		slide.addShape(content123w);

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();
		content123.setText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + ".");

		content123.setAnchor(new java.awt.Rectangle(10, 495, 560, 30));
		content123.setLineColor(new Color(203, 201, 176));
		content123.setLineWidth(2);

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(2.);
		contentP123.setSpaceBefore(10.);
		contentP123.setLineSpacing(110.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(15.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);

	}

	public void slideTrendTweetperHour(HSLFSlideShow ppt, HSLFSlide slide) throws IOException {
		setBackground(slide, version);
		judul(slide, "Trend Tweet per hour", 1);

		// ------------- add picture -----------//
		HSLFPictureData pd3a = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\TimeChartTWAccount.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3a = new HSLFPictureShape(pd3a);
		pictNew3a.setAnchor(new java.awt.Rectangle(270, 90, 610, 300));
		slide.addShape(pictNew3a);

		// ============content2===========//
		HSLFTextBox content12 = new HSLFTextBox();
		content12.setText("Total Tweet : " + " tweets");
		content12.appendText("\nTimeframe : " + " hours", false);

		content12.setAnchor(new java.awt.Rectangle(270, 430, 240, 55));
		content12.setLineColor(new Color(191, 200, 196));
		content12.setLineWidth(2);

		HSLFTextParagraph contentP12 = content12.getTextParagraphs().get(0);
		contentP12.setAlignment(TextAlign.JUSTIFY);
		contentP12.setSpaceAfter(2.);
		contentP12.setSpaceBefore(10.);
		contentP12.setLineSpacing(110.0);

		HSLFTextRun runContent12 = contentP12.getTextRuns().get(0);
		runContent12.setFontSize(15.);
		runContent12.setFontFamily("Calibri");
		runContent12.setBold(true);
		setTextColor(runContent12, version);
		slide.addShape(content12);

		// ============content3===========//
		HSLFTextBox content123w = new HSLFTextBox();
		content123w.setText("Max tweet= " + "tweets/hour");
		content123w.appendText("\nMin tweet= " + "tweets/hour", false);
		content123w.appendText("\nAv. tweet= " + "tweets/hour", false);
		content123w.appendText("\nDalam i jam, " + tema + " biasanya mengirimkan 1 tweet.", false);
		content123w.setAnchor(new java.awt.Rectangle(5, 70, 235, 140));
		content123w.setLineColor(new Color(150, 150, 150));
		content123w.setLineWidth(2);
		content123w.setFillColor(Color.WHITE);

		HSLFTextParagraph contentP123w = content123w.getTextParagraphs().get(0);
		contentP123w.setAlignment(TextAlign.JUSTIFY);
		contentP123w.setSpaceAfter(2.);
		contentP123w.setSpaceBefore(10.);
		contentP123w.setLineSpacing(110.0);

		HSLFTextRun runContent123w = contentP123w.getTextRuns().get(0);
		runContent123w.setFontSize(16.);
		runContent123w.setFontFamily("Corbel (Body)");
		setTextColor(runContent123w, version);
		slide.addShape(content123w);

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();
		content123.setText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + ".");

		content123.setAnchor(new java.awt.Rectangle(10, 495, 560, 30));
		content123.setLineColor(new Color(203, 201, 176));
		content123.setLineWidth(2);

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(2.);
		contentP123.setSpaceBefore(10.);
		contentP123.setLineSpacing(110.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(15.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);

	}

	public void slideOptimalTimeDiagram(HSLFSlideShow ppt, HSLFSlide slide) throws IOException {
		setBackground(slide, version);
		judul(slide, "Optimal Time Diagram", 1);

		// ------------- add picture -----------//
		HSLFPictureData pd3a = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\TimeChartTWAccount.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3a = new HSLFPictureShape(pd3a);
		pictNew3a.setAnchor(new java.awt.Rectangle(270, 90, 610, 300));
		slide.addShape(pictNew3a);

		// ============content2===========//
		HSLFTextBox content12 = new HSLFTextBox();
		content12.setText(
				"Gradasi warna berdasarkan frekuensi tweet pada hari dan jam. Semakin gelap, semakin tinggi frekuensi tweet akun @ShamsiAli2.");

		content12.setAnchor(new java.awt.Rectangle(270, 430, 640, 55));

		HSLFTextParagraph contentP12 = content12.getTextParagraphs().get(0);
		contentP12.setAlignment(TextAlign.JUSTIFY);
		contentP12.setSpaceAfter(2.);
		contentP12.setSpaceBefore(10.);
		contentP12.setLineSpacing(110.0);

		HSLFTextRun runContent12 = contentP12.getTextRuns().get(0);
		runContent12.setFontSize(16.);
		runContent12.setFontFamily("Corbel");
		setTextColor(runContent12, version);
		slide.addShape(content12);

		// ============content3===========//
		HSLFTextBox content123w = new HSLFTextBox();
		content123w.setText("Akun " + tema + " paling aktif pada Hari jumat, sekitar pukul 08.00 wib");

		content123w.setAnchor(new java.awt.Rectangle(5, 70, 235, 80));
		content123w.setLineColor(new Color(191, 200, 196));
		content123w.setLineWidth(3);
		content123w.setFillColor(Color.WHITE);

		HSLFTextParagraph contentP123w = content123w.getTextParagraphs().get(0);
		contentP123w.setAlignment(TextAlign.JUSTIFY);
		contentP123w.setSpaceAfter(2.);
		contentP123w.setSpaceBefore(10.);
		contentP123w.setLineSpacing(110.0);

		HSLFTextRun runContent123w = contentP123w.getTextRuns().get(0);
		runContent123w.setFontSize(16.);
		runContent123w.setFontFamily("Corbel (Body)");
		setTextColor(runContent123w, version);
		slide.addShape(content123w);

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();
		content123.setText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + ".");

		content123.setAnchor(new java.awt.Rectangle(10, 495, 560, 30));
		content123.setLineColor(new Color(203, 201, 176));
		content123.setLineWidth(2);

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(2.);
		contentP123.setSpaceBefore(10.);
		contentP123.setLineSpacing(110.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(15.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);
	}

	public void slideTopikPerbicaraan(HSLFSlide slide) {
		setBackground(slide, version);
		HSLFTextBox title = new HSLFTextBox();
		title.setText("Topik Pembicaraan");
		title.setAnchor(new java.awt.Rectangle(50, 20, 860, 100));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(36.);
		runTitle.setFontFamily("Corbel (Headings)");

		slide.addShape(title);

		HSLFTextBox contentx12 = new HSLFTextBox();
		contentx12.setAnchor(new java.awt.Rectangle(920, 50, 40, 440));
		contentx12.setFillColor(new Color(228, 228, 228));
		slide.addShape(contentx12);

		// -------------------------Content2------------------------//

		String username = "Ridwan Kamil";
		String hashtag_terbesar = "#kopasus";
		String jml_hashtag_terbesar = "200";

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("kata yng opaling banyak disebutkan oleh akun " + tema + " adalah ");
		content1.appendText(
				"Namanya sendiri, yaitu “shamsi” dan “ali”. #July4inJKT digunakan dalam tweet @usembassyjkt tentang peringatan hari kemerdekaan USA, yang kemudian di-retweet oleh @ShamsiAli2.",
				false);
		content1.appendText(
				"Beberapa tweetnya menceritakan pembangunan pesantren pertama di Amerika oleh komunitas Muslim di AS.",
				false);
		content1.setAnchor(new java.awt.Rectangle(10, 370, 900, 100));
		content1.setLineColor(new Color(150, 150, 150));
		content1.setLineWidth(3);

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.JUSTIFY);
		contentP1.setSpaceAfter(0.);
		contentP1.setSpaceBefore(0.);
		contentP1.setLineSpacing(100.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(18.);
		runContent1.setFontFamily("Corbel (Body)");
		setTextColor(runContent1, version);
		slide.addShape(content1);

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();
		content123.setText("Ket: Besar tulisan menunjukan tingkat frekuensi kata/hashtag tersebut di tweet");
		content123.appendText("\nKet: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + ".",
				false);

		content123.setAnchor(new java.awt.Rectangle(10, 500, 700, 30));

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(0.);
		contentP123.setSpaceBefore(0.);
		contentP123.setLineSpacing(100.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(14.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);

	}

	public void slideTweetsType(HSLFSlideShow ppt, HSLFSlide slide) throws IOException {
		setBackground(slide, version);
		judul(slide, "Tweet's Type", 1);

		// ------------- add picture -----------//
		HSLFPictureData pd3a = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\TimeChartTWAccount.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3a = new HSLFPictureShape(pd3a);
		pictNew3a.setAnchor(new java.awt.Rectangle(350, 80, 560, 350));
		slide.addShape(pictNew3a);

		// ============content3===========//
		HSLFTextBox content123w = new HSLFTextBox();
		content123w.setText("Berdasarkan data tweet " + tema + " dalam 33 hari terakhir, diketahui bahwa : ");
		content123w.appendText("\n64,1% dari total tweetnya merupakan asli tweet.", false);
		content123w.setAnchor(new java.awt.Rectangle(5, 70, 235, 200));

		HSLFTextParagraph contentP123w = content123w.getTextParagraphs().get(0);
		contentP123w.setAlignment(TextAlign.JUSTIFY);
		contentP123w.setSpaceAfter(2.);
		contentP123w.setSpaceBefore(10.);
		contentP123w.setLineSpacing(110.0);

		HSLFTextRun runContent123w = contentP123w.getTextRuns().get(0);
		runContent123w.setFontSize(16.);
		runContent123w.setFontFamily("Corbel (Body)");
		setTextColor(runContent123w, version);
		slide.addShape(content123w);

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();
		content123.setText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + ".");

		content123.setAnchor(new java.awt.Rectangle(10, 495, 560, 30));
		content123.setLineColor(new Color(203, 201, 176));
		content123.setLineWidth(2);

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(2.);
		contentP123.setSpaceBefore(10.);
		contentP123.setLineSpacing(110.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(15.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);

	}

	public void slideTweetsResponse(HSLFSlideShow ppt, HSLFSlide slide) throws IOException {
		setBackground(slide, version);
		judul(slide, "Tweet's Response", 1);

		// ------------- add picture -----------//
		HSLFPictureData pd3a = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\TimeChartTWAccount.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3a = new HSLFPictureShape(pd3a);
		pictNew3a.setAnchor(new java.awt.Rectangle(350, 80, 560, 350));
		slide.addShape(pictNew3a);

		// ============content3===========//
		HSLFTextBox content123w = new HSLFTextBox();
		content123w.setText("Berdasarkan data tweet " + tema + " dalam 33 hari terakhir, diketahui bahwa : ");
		content123w.appendText(
				"\n18,8% tweetnya telah di-retweet akun lain, dengan jumlah total akun yang me-retweet  44.255 akun.",
				false);
		content123w.setAnchor(new java.awt.Rectangle(5, 70, 235, 200));

		HSLFTextParagraph contentP123w = content123w.getTextParagraphs().get(0);
		contentP123w.setAlignment(TextAlign.JUSTIFY);
		contentP123w.setSpaceAfter(2.);
		contentP123w.setSpaceBefore(10.);
		contentP123w.setLineSpacing(110.0);

		HSLFTextRun runContent123w = contentP123w.getTextRuns().get(0);
		runContent123w.setFontSize(16.);
		runContent123w.setFontFamily("Corbel (Body)");
		setTextColor(runContent123w, version);
		slide.addShape(content123w);

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();
		content123.setText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + ".");

		content123.setAnchor(new java.awt.Rectangle(10, 495, 560, 30));
		content123.setLineColor(new Color(203, 201, 176));
		content123.setLineWidth(2);

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(2.);
		contentP123.setSpaceBefore(10.);
		contentP123.setLineSpacing(110.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(15.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);

	}

	public void slideTopEngagedAccounts(HSLFSlide slide) {
		setBackground(slide, version);
		HSLFTextBox title = new HSLFTextBox();
		title.setText("Top Engaged Accounts");
		title.setAnchor(new java.awt.Rectangle(50, 20, 860, 100));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(36.);
		runTitle.setFontFamily("Corbel (Headings)");

		slide.addShape(title);

		HSLFTextBox contentx12 = new HSLFTextBox();
		contentx12.setAnchor(new java.awt.Rectangle(920, 50, 40, 440));
		contentx12.setFillColor(new Color(228, 228, 228));
		slide.addShape(contentx12);

		// -------------------------Content2------------------------//

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("kata yng opaling banyak disebutkan oleh akun " + tema + " adalah ");
		content1.appendText(
				"Namanya sendiri, yaitu “shamsi” dan “ali”. #July4inJKT digunakan dalam tweet @usembassyjkt tentang peringatan hari kemerdekaan USA, yang kemudian di-retweet oleh @ShamsiAli2.",
				false);
		content1.appendText(
				"Beberapa tweetnya menceritakan pembangunan pesantren pertama di Amerika oleh komunitas Muslim di AS.",
				false);
		content1.setAnchor(new java.awt.Rectangle(10, 370, 900, 100));
		content1.setLineColor(new Color(150, 150, 150));
		content1.setLineWidth(3);

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.JUSTIFY);
		contentP1.setSpaceAfter(0.);
		contentP1.setSpaceBefore(0.);
		contentP1.setLineSpacing(100.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(18.);
		runContent1.setFontFamily("Corbel (Body)");
		setTextColor(runContent1, version);
		slide.addShape(content1);

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();
		content123.setText("Ket: Besar tulisan menunjukan tingkat frekuensi kata/hashtag tersebut di tweet");
		content123.appendText("\nKet: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + ".",
				false);

		content123.setAnchor(new java.awt.Rectangle(10, 500, 700, 30));

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(0.);
		contentP123.setSpaceBefore(0.);
		contentP123.setLineSpacing(100.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(14.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);
	}

	public void slideTopSharedLinks(HSLFSlide slide) {
		setBackground(slide, version);
		HSLFTextBox title = new HSLFTextBox();
		title.setText("Top Shared Links");
		title.setAnchor(new java.awt.Rectangle(50, 20, 860, 100));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(36.);
		runTitle.setFontFamily("Corbel (Headings)");

		slide.addShape(title);

		HSLFTextBox contentx12 = new HSLFTextBox();
		contentx12.setAnchor(new java.awt.Rectangle(920, 50, 40, 440));
		contentx12.setFillColor(new Color(228, 228, 228));
		slide.addShape(contentx12);

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();

		content123.appendText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + ".", false);

		content123.setAnchor(new java.awt.Rectangle(10, 505, 700, 30));

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(0.);
		contentP123.setSpaceBefore(0.);
		contentP123.setLineSpacing(100.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(15.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);

	}

	public void slideTopMentionAccount(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Top Mention Account", 3);
		// -------------------------Content2------------------------//

		String username = "Ridwan Kamil";
		String influencer = "@bts.bighitofficial";

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Grafik berikut menggambarkan 5 akun yang paling sering disebut oleh akun " + username);

		content1.appendText(". Akun yang paling sering di-mention oleh akun " + username + " adalah " + influencer
				+ ". Akun ini merupakan salah satu akun influencer.", false);

		content1.setAnchor(new java.awt.Rectangle(120, 400, 760, 100));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.CENTER);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(20.);
		runContent1.setFontFamily("SimSun-ExtB (Body)");
		setTextColor(runContent1, version);
		slide.addShape(content1);

	}

	public void slideTopEngagedAccount2(HSLFSlide slide) {
		setBackground(slide, version);
		HSLFTextBox title = new HSLFTextBox();
		title.setText("Top Engaged Account");
		title.setAnchor(new java.awt.Rectangle(50, 20, 860, 100));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(36.);
		runTitle.setFontFamily("Corbel (Headings)");

		slide.addShape(title);

		HSLFTextBox contentx12 = new HSLFTextBox();
		contentx12.setAnchor(new java.awt.Rectangle(920, 50, 40, 440));
		contentx12.setFillColor(new Color(228, 228, 228));
		slide.addShape(contentx12);

		// -------------------------Content2------------------------//

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText(
				"Akun-akun berikut adalah akun-akun yang dalam aktivitasnya me-mention/me-retweet @ShamsiAli2.");

		content1.setAnchor(new java.awt.Rectangle(70, 370, 900, 100));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.JUSTIFY);
		contentP1.setSpaceAfter(0.);
		contentP1.setSpaceBefore(0.);
		contentP1.setLineSpacing(100.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(18.);
		runContent1.setFontFamily("Corbel (Body)");
		setTextColor(runContent1, version);
		slide.addShape(content1);

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();

		content123.appendText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + ".", false);

		content123.setAnchor(new java.awt.Rectangle(10, 500, 700, 30));

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(0.);
		contentP123.setSpaceBefore(0.);
		contentP123.setLineSpacing(100.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(15.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);

	}

	public void slideTopTweets(HSLFSlide slide) {
		setBackground(slide, version);
		HSLFTextBox title = new HSLFTextBox();
		title.setText("Top Tweets");
		title.setAnchor(new java.awt.Rectangle(50, 20, 860, 100));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(36.);
		runTitle.setFontFamily("Corbel (Headings)");

		slide.addShape(title);

		HSLFTextBox contentx12 = new HSLFTextBox();
		contentx12.setAnchor(new java.awt.Rectangle(920, 50, 40, 440));
		contentx12.setFillColor(new Color(228, 228, 228));
		slide.addShape(contentx12);

		// -------------------------Content2------------------------//

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Tabel di atas adalah urutan tweet dari akun @ShamsiAli2 yang paling banyak di-retweet.");

		content1.setAnchor(new java.awt.Rectangle(70, 370, 900, 100));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.JUSTIFY);
		contentP1.setSpaceAfter(0.);
		contentP1.setSpaceBefore(0.);
		contentP1.setLineSpacing(100.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(18.);
		runContent1.setFontFamily("Corbel (Body)");
		setTextColor(runContent1, version);
		slide.addShape(content1);

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();

		content123.appendText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + ".", false);

		content123.setAnchor(new java.awt.Rectangle(10, 500, 700, 30));

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(0.);
		contentP123.setSpaceBefore(0.);
		contentP123.setLineSpacing(100.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(15.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);

	}

	public void slideNetworkofRetweet(HSLFSlide slide) {
		setBackground(slide, 1);
		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network by Page");
		title.setAnchor(new java.awt.Rectangle(50, 20, 860, 100));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(36.);
		runTitle.setFontFamily("Corbel (Headings)");
		runTitle.setFontColor(Color.white);

	}

	public void slideConclusion(HSLFSlide slide) {
		setBackground(slide, 2);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Kesimpulan");
		title.setAnchor(new java.awt.Rectangle(50, 30, 860, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(48.);
		runTitle.setFontFamily("Times New Roman");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, 2);

		slide.addShape(title);

	}
}
