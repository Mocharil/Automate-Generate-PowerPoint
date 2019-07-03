package com.ebdesk.report.slide;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

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
import org.bson.BSONObject;
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
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.DBObject;
import com.mongodb.Mongo;
import com.mongodb.MongoClient;
import com.mongodb.MongoCredential;
import com.mongodb.ServerAddress;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoCursor;
import com.mongodb.client.MongoDatabase;

public class SNAInstagramReport {
	private int version;
	private String filename;
	private int width;
	private int height;
	private static String mongoHost = "192.168.114.162";
	private static String mongoPort = "27017";
	private static String mongoUsername = "dev";
	private static String mongoPassword = "Rahas!adev20!8";
	private static String mongoSource = "admin";
	private String tema;
	private static String db;
	private static String collect;

	public SNAInstagramReport(String filename, int version, int width, int height, String tema, String db,
			String collect) {
		this.filename = filename;
		this.version = version;
		this.height = height;
		this.width = width;
		this.tema = tema;
		this.db = db;
		this.collect = collect;
	}

	public void execute() throws Exception {
		// -----------------MongoDB--------------------//
		final MongoCredential credential = MongoCredential.createScramSha1Credential(mongoUsername, mongoSource,
				mongoPassword.toCharArray());
		ServerAddress serverAddress = new ServerAddress(mongoHost, Integer.parseInt(mongoPort));
		MongoClient mongoClient = new MongoClient(serverAddress, new ArrayList<MongoCredential>() {
			/**
			 * 
			 */
			private static final long serialVersionUID = 1L;

			{
				add(credential);
			}
		});

		MongoDatabase database = mongoClient.getDatabase(db);
		MongoCollection<Document> collection = database.getCollection(collect);

		MongoCursor<Document> cursor = collection.find().iterator();
		ArrayList<Document> list = new ArrayList<Document>();

		try {
			while (cursor.hasNext()) {
				list.add(cursor.next());

			}
		} catch (Exception e) {
			System.out.println("gagal");
		} finally {
			cursor.close();
		}

		String g1;
		String[] time1;
		String[] time2;
		String[] bulan = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
		int nilai_bulan = 1;

		int tanggal1 = 1;
		int tanggal_awal = 999999999;
		int index1 = 0;
		int index2 = 0;
		int tanggal_akhir = 0;
		float image = 0;
		float MM = 0;
		float video = 0;

		Map<String, Integer> map = new HashMap<String, Integer>();

		ArrayList<String> settanggal = new ArrayList<String>();
		ArrayList<String> listhtg = new ArrayList<String>();
		ArrayList<String> listnama = new ArrayList<String>();

		List<String> myString = new ArrayList<String>();

		for (int ii = 0; ii < list.size(); ii++) {
			// =====================TEXT====================//
			ArrayList x7 = (ArrayList) list.get(ii).get("hashtag"); // hashtag
			try {
				if (x7.size() != 0) {
					for (int ii1 = 0; ii1 < x7.size(); ii1++) {
						listhtg.add("#" + String.valueOf(x7.get(ii1))); // hashtag
					}
				}
			} catch (Exception e) {
			}

			ArrayList y7 = (ArrayList) list.get(ii).get("hahstag"); // hashtag
			try {
				if (y7.size() != 0) {
					for (int ii1 = 0; ii1 < y7.size(); ii1++) {
						listhtg.add("#" + String.valueOf(y7.get(ii1))); // hashtag
					}
				}
			} catch (Exception e) {
			}

			Pattern rex = Pattern.compile("(?<!\\S)\\p{Alpha}+(?!\\S)"); // regex words

			try {
				Matcher mx = rex.matcher((String) list.get(ii).get("text"));
				while (mx.find()) {
					myString.add(mx.group(0)); // wordcloud
				}
			} catch (Exception e) {

			}

			// =========USERNAME=============//
			listnama.add("@" + String.valueOf(list.get(ii).get("username")));

			// ===============TIME==============//
			g1 = String.valueOf(list.get(ii).get("created_at").toString());
			time1 = g1.split("\\s");
			for (int bul = 0; bul < 12; bul++) {

				if (String.valueOf(time1[1]).equals(String.valueOf(bulan[bul]))) {

					nilai_bulan = (bul + 1) * 100;
					tanggal1 = nilai_bulan + Integer.parseInt(time1[2]) + Integer.parseInt(time1[5]) * 10000;

				}
			}

			settanggal.add(String.valueOf(time1[2]) + "/" + String.valueOf(time1[1]) + "/" + String.valueOf(time1[5]));
			map.put(String.valueOf(time1[2]) + "/" + String.valueOf(time1[1]) + "/" + String.valueOf(time1[5]), ii);

			if (tanggal_awal > tanggal1) {
				tanggal_awal = tanggal1;
				index1 = ii;
			}

			if (tanggal_akhir < tanggal1) {
				tanggal_akhir = tanggal1;
				index2 = ii;
			}

			// ==================type post================//
			if (String.valueOf(list.get(ii).get("type")).equals("image")) {
				image += 1;
			} else if (String.valueOf(list.get(ii).get("type")).equals("carousel")) {
				MM += 1;
			} else if (String.valueOf(list.get(ii).get("type")).equals("video")) {
				video += 1;
			}

		}

		// ================tanggal=================//
		String ga = (String) list.get(index1).get("created_at").toString();
		time1 = ga.split("\\s"); // tanggal awal
		String gk = (String) list.get(index2).get("created_at").toString();
		time2 = gk.split("\\s"); // tanggal akhir

		HashSet<String> settime1 = new HashSet<String>(settanggal);

		// ===============TIME SERIES==============//
		String[] time3;
		final TimeSeries series = new TimeSeries("Random Data");
		int indexmap;

		Object[] r = settime1.toArray();
		for (int oo = 0; oo < settime1.size(); oo++) {
			try {
				indexmap = map.get(String.valueOf(r[oo]));
				g1 = (String) list.get(indexmap).get("created_at").toString();
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
		JFreeChart timechart = ChartFactory.createTimeSeriesChart(null, "Time", "Instagram Count", dataset, false,
				false, false);
		timechart.setBackgroundPaint(Color.white);
		XYPlot plot = (XYPlot) timechart.getPlot();
		plot.setBackgroundPaint(Color.white);
		int width1 = 580; /* Width of the image */
		int height1 = 370; /* Height of the image */
		File timeChart = new File("TimeChartIGv1.png");
		ChartUtilities.saveChartAsPNG(timeChart, timechart, width1, height1);

		// ================grafik type post==============//
		float persen_video = video * 100 / list.size();
		float persen_MM = MM * 100 / list.size();
		float persen_image = image * 100 / list.size();
		float persentase_tertinggi = persen_video;
		String jenis_post = "Video";

		if (persentase_tertinggi < persen_MM) {
			persentase_tertinggi = persen_MM;
			jenis_post = "Multiple Media";
		}
		if (persentase_tertinggi < persen_image) {
			jenis_post = "Photo";
			persentase_tertinggi = persen_image;
		}

		PieChartReport chart1 = new PieChartReport();
		chart1.ringchart("Photo", persen_image, "Video", persen_video, "Multiple Media", persen_MM, null, 0, "pieIGv1.png", Color.white);

		// ===========BARCHART===========//
		BarChart chart7 = new BarChart();
		ArrayList<String> ranktgl = chart7.bar(settanggal, settime1, null, null, 1, 0, null);
		ArrayList<String> ranknama = chart7.bar(listnama, new HashSet<String>(listnama), "barChartIGv1.png", null, 5, 2,
				"red");
		ArrayList<String> rankhtg = chart7.bar(listhtg, new HashSet<String>(listhtg), "barcharthtgIGv1.png", null, 10,
				1, "red");

		// ===============DATA=============//
		int jml_htg = Collections.frequency(listhtg, rankhtg.get(0));
		int jml_nama = Collections.frequency(listnama, ranknama.get(0));
		int post_puncak = Collections.frequency(settanggal, String.valueOf(ranktgl.get(0)));

		// ===========wordcloud===========//
		List<String> myStringhtg = new ArrayList<String>(listhtg);
		Wordcloud word = new Wordcloud(); // wordcloud
		List<WordFrequency> wc = word.readCNN("wcIGv1.png", myString, "blue");
		List<WordFrequency> wchtg = word.readCNN("wchtgIGv1.png", myStringhtg, "blue");

		// ====================SLIDE================//

		HSLFSlideShow ppt = new HSLFSlideShow();
		ppt.setPageSize(new java.awt.Dimension(width, height));

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\ig_logo.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(25, 504, 30, 30));

		// Slide 1
		HSLFSlide slide1 = ppt.createSlide();
		slideMaster(slide1, time2);
		slide1.addShape(pictNew);

		// Slide 2
		HSLFSlide slide2 = ppt.createSlide();
		slideInformation(slide2, time1, time2);
		slide2.addShape(pictNew);

		// Slide 3
		HSLFSlide slide3 = ppt.createSlide();
		slideTrendPost1(ppt, slide3, time1, time2, ranktgl, post_puncak, list.size(), wc, wchtg);
		slide3.addShape(pictNew);

		// Slide 4
		HSLFSlide slide4 = ppt.createSlide();
		slidePostType(ppt, slide4, persentase_tertinggi, jenis_post, list.size());
		slide4.addShape(pictNew);

		// Slide 5
		HSLFSlide slide5 = ppt.createSlide();
		slide5akun(ppt, slide5, ranknama, jml_nama);
		slide5.addShape(pictNew);

		// // Slide 6
		HSLFSlide slide6 = ppt.createSlide();
		slideNetworkHashtag(slide6);
		slide6.addShape(pictNew);

		// // Slide 7
		HSLFSlide slide7 = ppt.createSlide();
		slideNetworkHashtag2(slide7);
		slide7.addShape(pictNew);

		// ------------- add picturegephi -----------//
		HSLFPictureData pd3gx = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\gephi.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3gx = new HSLFPictureShape(pd3gx);
		pictNew3gx.setAnchor(new java.awt.Rectangle(30, 60, 460, 430));
		slide7.addShape(pictNew3gx);

		// Slide 8
		HSLFSlide slide8 = ppt.createSlide();
		slide10hashtag(ppt, slide8, rankhtg, jml_htg);
		slide8.addShape(pictNew);

		// // Slide 9
		HSLFSlide slide9 = ppt.createSlide();
		slideNetworkHashtag3(slide9);
		slide9.addShape(pictNew);

		// ------------- add picturegephi -----------//
		HSLFPictureData pd3x = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\gephi.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3x = new HSLFPictureShape(pd3x);
		pictNew3x.setAnchor(new java.awt.Rectangle(30, 60, 460, 430));
		slide9.addShape(pictNew3x);

		// // Slide 10
		HSLFSlide slide10 = ppt.createSlide();
		slideNetworkHashtag4(slide10);
		slide10.addShape(pictNew);

		// ------------- add picturegephi -----------//
		HSLFPictureData pd3gx11 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\gephi.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3gx11 = new HSLFPictureShape(pd3gx11);
		pictNew3gx11.setAnchor(new java.awt.Rectangle(30, 60, 460, 430));
		slide10.addShape(pictNew3gx11);

		// // Slide 11
		HSLFSlide slide11 = ppt.createSlide();
		slideNetworkbyMention1(slide11);
		slide11.addShape(pictNew);

		// ------------- add picturegephi -----------//
		HSLFPictureData pd3gx2 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\gephi.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3gx2 = new HSLFPictureShape(pd3gx2);
		pictNew3gx2.setAnchor(new java.awt.Rectangle(30, 60, 460, 430));
		slide11.addShape(pictNew3gx2);

		// // Slide 12
		HSLFSlide slide12 = ppt.createSlide();
		slideNetworkbyMention2(slide12);
		slide12.addShape(pictNew);

		// ------------- add picturegephi -----------//
		HSLFPictureData pd3gx3 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\gephi.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3gx3 = new HSLFPictureShape(pd3gx3);
		pictNew3gx3.setAnchor(new java.awt.Rectangle(420, 60, 470, 440));
		slide12.addShape(pictNew3gx3);

		// // Slide 13
		HSLFSlide slide13 = ppt.createSlide();
		slideKesimpulan(slide13, time1, time2);
		slide13.addShape(pictNew);

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
			fill.setFillType(HSLFFill.FILL_SHADE);
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

	// =======content Platform========//
	public void Platform(HSLFSlide slide, int warna) {

		HSLFTextBox contentp = new HSLFTextBox();

		contentp.setText("Platform: Instagram");

		contentp.setAnchor(new java.awt.Rectangle(50, 510, 350, 50));

		HSLFTextParagraph titlePp = contentp.getTextParagraphs().get(0);
		titlePp.setAlignment(TextAlign.JUSTIFY);
		titlePp.setSpaceAfter(0.);
		titlePp.setSpaceBefore(0.);
		titlePp.setLineSpacing(110.0);

		HSLFTextRun runTitlep = titlePp.getTextRuns().get(0);
		runTitlep.setFontSize(15.);
		runTitlep.setFontFamily("Times New Roman");
		runTitlep.setBold(true);
		runTitlep.setItalic(true);
		if (warna == 1) {
			runTitlep.setFontColor(new Color(227, 52, 103));
		} else {
			runTitlep.setFontColor(new Color(248, 208, 219));
		}
		slide.addShape(contentp);
	}

	public void slideMaster(HSLFSlide slide, String[] time2) {
		setBackground(slide, 1);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("SOCIAL");
		title.appendText("NETWORK", true);
		title.appendText("ANALYSIS", true);
		title.setAnchor(new java.awt.Rectangle(50, 80, 500, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.LEFT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);
		titleP.setLineSpacing(80.0);
		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(80.);
		runTitle.setFontFamily("Times New Roman");
		runTitle.setBold(true);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		runTitle.setFontColor(Color.red);

		HSLFTextParagraph titleP1 = title.getTextParagraphs().get(2);
		HSLFTextRun runTitle1 = titleP1.getTextRuns().get(0);
		runTitle1.setFontSize(80.);
		runTitle1.setFontFamily("Times New Roman");
		runTitle1.setBold(true);
		runTitle1.setItalic(true);
		runTitle1.setUnderlined(false);
		setTextColor(runTitle1, 1);
		slide.addShape(title);

		// --------------------Investigation--------------------
		HSLFTextBox issue = new HSLFTextBox();
		issue.setText("INVESTIGASI ISU " + tema);
		issue.setAnchor(new java.awt.Rectangle(50, 350, 900, 50));
		HSLFTextParagraph issueP = issue.getTextParagraphs().get(0);
		HSLFTextRun runIssue = issueP.getTextRuns().get(0);
		runIssue.setFontSize(30.);
		runIssue.setFontFamily("Times New Roman");
		runIssue.setBold(true);
		runIssue.setItalic(true);
		setTextColor(runIssue, 1);
		slide.addShape(issue);

		// -------------------------Date------------------------
		HSLFTextBox contentpt = new HSLFTextBox();
		contentpt.setText(time2[2] + "-" + time2[1] + "-" + time2[5]);

		contentpt.setAnchor(new java.awt.Rectangle(50, 430, 350, 50));

		HSLFTextParagraph titlePpt = contentpt.getTextParagraphs().get(0);
		titlePpt.setAlignment(TextAlign.JUSTIFY);
		titlePpt.setSpaceAfter(0.);
		titlePpt.setSpaceBefore(0.);
		titlePpt.setLineSpacing(110.0);

		HSLFTextRun runTitlept = titlePpt.getTextRuns().get(0);
		runTitlept.setFontSize(15.);
		runTitlept.setFontFamily("Times New Roman");
		runTitlept.setBold(false);
		runTitlept.setItalic(true);
		setTextColor(runTitlept, 1);

		slide.addShape(contentpt);

		// =======content Platform========//
		Platform(slide, 2);

	}

	public void slideInformation(HSLFSlide slide, String[] time1, String[] time2) {
		setBackground(slide, 2);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Information");
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

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();

		content.setText("Dalam kurun " + time1[2] + "/" + time1[1] + "/" + time1[5]);
		content.appendText(" - ", false);
		content.appendText(time2[2] + "/" + time2[1] + "/" + time2[5], false);
		content.appendText(", issue mengenai '", false);
		content.appendText(tema, false);
		content.appendText("' sedang ramai diperbincangkan di media social Instagram.", false);
		content.setAnchor(new java.awt.Rectangle(500, 165, 400, 130));

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(22.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, 2);

		slide.addShape(content);

		// =======content Platform========//
		Platform(slide, 1);
	}

	public void slideTrendPost1(HSLFSlideShow ppt, HSLFSlide slide, String[] time1, String[] time2,
			ArrayList<String> ranktgl, int post_puncak, int jml_data, List<WordFrequency> rankhtg,
			List<WordFrequency> wchtg) throws IOException {

		setBackground(slide, 2);

		HSLFTextBox title = new HSLFTextBox();

		title.setText("Trend Post");
		title.appendText(tema, true);
		title.appendText(time2[2] + "/" + time2[1] + "/" + time2[5], true);

		title.setAnchor(new java.awt.Rectangle(23, 30, 860, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.JUSTIFY);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(38.);
		runTitle.setFontFamily("Times New Roman");
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, 2);

		slide.addShape(title);
		// ------------- add picture grafik -----------//
		HSLFPictureData pd3 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\TimeChartIGv1.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3 = new HSLFPictureShape(pd3);
		pictNew3.setAnchor(new java.awt.Rectangle(400, 280, 500, 250));
		slide.addShape(pictNew3);

		// ------------- add picture hashtag -----------//
		HSLFPictureData pd1 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\wchtgIGv1.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
		pictNew1.setAnchor(new java.awt.Rectangle(420, 60, 190, 190));
		slide.addShape(pictNew1);

		// ------------- add picture wordcloud -----------//
		HSLFPictureData pd2 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\wcIGv1.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew2 = new HSLFPictureShape(pd2);
		pictNew2.setAnchor(new java.awt.Rectangle(690, 60, 190, 190));
		slide.addShape(pictNew2);

		// =======content2========//
		HSLFTextBox content2 = new HSLFTextBox();

		content2.setText("\nTrend Postingan dari tanggal " + time1[2] + "/" + time1[1] + "/" + time1[5]);
		content2.appendText(" - ", false);
		content2.appendText(time2[2] + "/" + time2[1] + "/" + time2[5] + " : ", false);

		content2.appendText("1.  Puncak pembicaraan tertinggi dari issue “", true);
		content2.appendText(tema, false);
		content2.appendText("“ terjadi pada tanggal ", false);
		content2.appendText(ranktgl.get(0), false);
		content2.appendText(". Jumlah post pada tanggal ini adalah ", false);
		content2.appendText(String.valueOf(post_puncak), false);
		content2.appendText("/", false);
		content2.appendText(String.valueOf(jml_data), false);
		content2.appendText(" post.", false);
		content2.appendText(" Rata-rata post selama kurun waktu tersebut sebesar ", false);
		content2.appendText(String.valueOf(post_puncak / 24), false);
		content2.appendText(" post/hour.", false);

		content2.appendText("\n2.  Hashtag yang sering dibicarakan dalam issue ini adalah : ", true);
		content2.appendText(wchtg.get(0).getWord() + ", " + wchtg.get(1).getWord() + " dan " + wchtg.get(2).getWord(),
				false);
		content2.appendText(".\nKata – kata yang sering dibicarakan dalam caption pada tanggal ini adalah : ", false);
		content2.appendText(
				rankhtg.get(0).getWord() + ", " + rankhtg.get(1).getWord() + " dan " + rankhtg.get(2).getWord() + ".",
				false);
		content2.appendText("\n\n\t\t       Jumlah data: ", true);
		content2.appendText(String.valueOf(jml_data), false);
		content2.appendText(" post", false);

		content2.setAnchor(new java.awt.Rectangle(10, 170, 353, 300));

		HSLFTextParagraph titleP2 = content2.getTextParagraphs().get(0);
		titleP2.setAlignment(TextAlign.JUSTIFY);
		titleP2.setSpaceAfter(0.);
		titleP2.setSpaceBefore(0.);
		titleP2.setLineSpacing(110.0);

		HSLFTextRun runTitle2 = titleP2.getTextRuns().get(0);
		runTitle2.setFontSize(15.);
		runTitle2.setFontFamily("Calibri");
		runTitle2.setBold(false);
		setTextColor(runTitle2, 2);

		slide.addShape(content2);

		// =======content 3========//
		HSLFTextBox content3 = new HSLFTextBox();

		content3.setText("Hashtag");

		content3.setAnchor(new java.awt.Rectangle(480, 250, 100, 100));

		HSLFTextParagraph titleP3 = content3.getTextParagraphs().get(0);
		titleP3.setAlignment(TextAlign.JUSTIFY);
		titleP3.setSpaceAfter(0.);
		titleP3.setSpaceBefore(0.);
		titleP3.setLineSpacing(110.0);

		HSLFTextRun runTitle3 = titleP3.getTextRuns().get(0);
		runTitle3.setFontSize(16.);
		runTitle3.setFontFamily("Calibri");
		runTitle3.setBold(false);
		setTextColor(runTitle3, 2);

		slide.addShape(content3);

		// =======content 4========//
		HSLFTextBox content4 = new HSLFTextBox();

		content4.setText("Caption");

		content4.setAnchor(new java.awt.Rectangle(760, 250, 100, 100));

		HSLFTextParagraph titleP4 = content4.getTextParagraphs().get(0);
		titleP4.setAlignment(TextAlign.JUSTIFY);
		titleP4.setSpaceAfter(0.);
		titleP4.setSpaceBefore(0.);
		titleP4.setLineSpacing(110.0);

		HSLFTextRun runTitle4 = titleP4.getTextRuns().get(0);
		runTitle4.setFontSize(16.);
		runTitle4.setFontFamily("Calibri");
		runTitle4.setBold(false);
		setTextColor(runTitle4, 2);

		slide.addShape(content4);

		// =======content Platform========//
		Platform(slide, 1);
	}

	public void slidePostType(HSLFSlideShow ppt, HSLFSlide slide, float jml_post, String jenis_post, int jml_data)
			throws IOException {
		setBackground(slide, 2);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Post Type");
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

		// --------------------grafik-------------------//
		HSLFPictureData pd4 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\pieIGv1.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew4 = new HSLFPictureShape(pd4);
		pictNew4.setAnchor(new java.awt.Rectangle(45, 65, 440, 390));
		slide.addShape(pictNew4);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();

		content.setText("Hasil diagram disamping menunjukkan bahwa mayoritas penyebaran issue dilakukan melalui ");
		content.appendText(jenis_post, false);
		content.appendText(" dengan persentase sebesar ", false);
		content.appendText(String.valueOf(jml_post), false);
		content.appendText(" %.", false);

		content.setAnchor(new java.awt.Rectangle(500, 165, 360, 350));

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(22.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, 2);

		slide.addShape(content);

		// ======================content2==================================
		HSLFTextBox content2 = new HSLFTextBox();
		content2.setText("Jumlah Data : " + jml_data + " Post");
		content2.setAnchor(new java.awt.Rectangle(50, 470, 400, 130));

		HSLFTextParagraph contentP2 = content2.getTextParagraphs().get(0);
		contentP2.setAlignment(TextAlign.JUSTIFY);
		contentP2.setSpaceAfter(0.);
		contentP2.setSpaceBefore(0.);
		contentP2.setLineSpacing(110.0);

		HSLFTextRun runContent2 = contentP2.getTextRuns().get(0);
		runContent2.setFontSize(20.);
		runContent2.setFontFamily("Calibri");
		runContent2.setBold(false);
		setTextColor(runContent2, 2);

		slide.addShape(content2);

		// =======content Platform========//
		Platform(slide, 1);
	}

	public void slide5akun(HSLFSlideShow ppt, HSLFSlide slide, ArrayList<String> ranknama, int jml_nama)
			throws IOException {

		setBackground(slide, 2);

		HSLFTextBox title = new HSLFTextBox();

		title.setText("\n\n5 Akun Yang\nPaling\nSering Memposting");

		title.setAnchor(new java.awt.Rectangle(30, 40, 860, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.JUSTIFY);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Times New Roman");
		runTitle.setBold(true);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, 2);

		slide.addShape(title);

		// ------------- add picture grafik -----------//
		HSLFPictureData pd5 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barChartIGv1.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew5 = new HSLFPictureShape(pd5);
		pictNew5.setAnchor(new java.awt.Rectangle(430, 120, 490, 320));
		slide.addShape(pictNew5);

		// ---------content2---------//

		HSLFTextBox content2 = new HSLFTextBox();

		content2.appendText("Akun yang paling sering memposting adalah ", false);
		content2.appendText(ranknama.get(0) + " dengan jumlah " + String.valueOf(jml_nama) + " kali digunakan.", false);
		content2.setAnchor(new java.awt.Rectangle(23, 380, 380, 50));

		HSLFTextParagraph titleP2 = content2.getTextParagraphs().get(0);
		titleP2.setAlignment(TextAlign.JUSTIFY);
		titleP2.setSpaceAfter(0.);
		titleP2.setSpaceBefore(0.);
		titleP2.setLineSpacing(110.0);

		HSLFTextRun runTitle2 = titleP2.getTextRuns().get(0);
		runTitle2.setFontSize(22.);
		runTitle2.setFontFamily("Calibri");
		runTitle2.setBold(false);
		setTextColor(runTitle2, 2);

		slide.addShape(content2);

		// =======content Platform========//
		Platform(slide, 1);
	}

	public void slideNetworkHashtag(HSLFSlide slide) {
		setBackground(slide, 1);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network\nHashtag");
		title.setAnchor(new java.awt.Rectangle(500, 10, 300, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.LEFT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(46.);
		runTitle.setFontFamily("Times New Roman");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, 1);

		slide.addShape(title);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();

		String issue = String.format("%s", " 'Bumi diserang Alien'");
		String persentase_modularity_tertinggi_1 = String.format("%s", "50");
		String persentase_modularity_tertinggi_2 = String.format("%s", "40");
		String persentase_modularity_tertinggi_3 = String.format("%s", "30");
		String persentase_modularity_tertinggi_4 = String.format("%s", "20");
		String persentase_modularity_tertinggi_5 = String.format("%s", "10");
		String jml_nodes = String.format("%s", "86");
		String jml_edges = String.format("%s", "78");
		String kelompok = String.format("%s", "91");
		String kelompok_besar = String.format("%s", "5");
		String persentase_modularity_terbesar = String.format("%s", "1");

		content.setText(
				"Network disamping merupakan network yang terdiri dari akun dan hashtag-hashtag yang digunakan dalam issue ");
		content.appendText(issue, false);
		content.appendText(". Network Hashtag terbagi menjadi ", false);
		content.appendText(kelompok, false);
		content.appendText(" kelompok. Berikut ", false);
		content.appendText(kelompok_besar, false);
		content.appendText(" Kelompok Terbesar: ", false);
		content.appendText(" ", true);
		content.appendText("• Kelompok 1	: ", true);
		content.appendText(persentase_modularity_tertinggi_1 + " %\n", false);
		content.appendText("• Kelompok 2	: ", true);
		content.appendText(persentase_modularity_tertinggi_2 + " %\n", false);
		content.appendText("• Kelompok 3	: ", true);
		content.appendText(persentase_modularity_tertinggi_3 + " %\n", false);
		content.appendText("• Kelompok 4	: ", true);
		content.appendText(persentase_modularity_tertinggi_4 + " %\n", false);
		content.appendText("• Kelompok 5	: ", true);
		content.appendText(persentase_modularity_tertinggi_5 + " %\n", false);

		String hashtag = String.format("%s", "54");
		String account = String.format("%s", "58");

		content.appendText(" ", true);
		content.appendText("Dari network hashtag, kelompok yang paling luas jaringannya adalah kelompok ", true);
		content.appendText(persentase_modularity_terbesar, false);

		content.appendText("Jumlah Nodes	: ", true);
		content.appendText(jml_nodes, false);
		content.appendText("Jumlah Edges: ", true);
		content.appendText(jml_edges, false);
		content.appendText("Hashtag	: ", true);
		content.appendText(hashtag, false);
		content.appendText("Account	: ", true);
		content.appendText(account, false);

		content.setAnchor(new java.awt.Rectangle(500, 130, 420, 130));

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(105.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(18.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, 1);

		slide.addShape(content);

		// =======content Platform========//
		Platform(slide, 2);
	}

	public void slideNetworkHashtag2(HSLFSlide slide) {
		setBackground(slide, 1);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network\nHashtag");
		title.setAnchor(new java.awt.Rectangle(500, 10, 300, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.LEFT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(48.);
		runTitle.setFontFamily("Times New Roman");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, 1);

		slide.addShape(title);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();
		String persentase_modularity_tertinggi_1 = String.format("%s", "50");
		String persentase_modularity_tertinggi_2 = String.format("%s", "40");
		String persentase_modularity_tertinggi_3 = String.format("%s", "30");
		String persentase_modularity_tertinggi_4 = String.format("%s", "20");
		String persentase_modularity_tertinggi_5 = String.format("%s", "10");
		String jml_nodes = String.format("%s", "86");
		String jml_edges = String.format("%s", "78");
		String hashtag_1_kelompok_5 = String.format("%s", "#aku24");
		String hashtag_2_kelompok_5 = String.format("%s", "#aku23");
		String hashtag_3_kelompok_5 = String.format("%s", "#aku55");

		content.setText("Hashtag-hashtag yang sering digunakan dalam setiap kelompok: ");
		content.appendText(" ", true);
		content.appendText("• Kelompok 1	: ", true);
		content.appendText(persentase_modularity_tertinggi_1 + " %\n", false);
		content.appendText(hashtag_1_kelompok_5 + ", " + hashtag_2_kelompok_5 + ", " + hashtag_3_kelompok_5, false);
		content.appendText("• Kelompok 2	: ", true);
		content.appendText(persentase_modularity_tertinggi_2 + " %\n", false);
		content.appendText(hashtag_1_kelompok_5 + ", " + hashtag_2_kelompok_5 + ", " + hashtag_3_kelompok_5, false);
		content.appendText("• Kelompok 3	: ", true);
		content.appendText(persentase_modularity_tertinggi_3 + " %\n", false);
		content.appendText(hashtag_1_kelompok_5 + ", " + hashtag_2_kelompok_5 + ", " + hashtag_3_kelompok_5, false);
		content.appendText("• Kelompok 4	: ", true);
		content.appendText(persentase_modularity_tertinggi_4 + " %\n", false);
		content.appendText(hashtag_1_kelompok_5 + ", " + hashtag_2_kelompok_5 + ", " + hashtag_3_kelompok_5, false);
		content.appendText("• Kelompok 5	: ", true);
		content.appendText(persentase_modularity_tertinggi_5 + " %\n", false);
		content.appendText(hashtag_1_kelompok_5 + ", " + hashtag_2_kelompok_5 + ", " + hashtag_3_kelompok_5, false);
		String hashtag = String.format("%s", "54");
		String account = String.format("%s", "58");

		content.appendText(" ", true);

		content.appendText("Jumlah Nodes	: ", true);
		content.appendText(jml_nodes, false);
		content.appendText("Jumlah Edges: ", true);
		content.appendText(jml_edges, false);
		content.appendText("Hashtag	: ", true);
		content.appendText(hashtag, false);
		content.appendText("Account	: ", true);
		content.appendText(account, false);

		content.setAnchor(new java.awt.Rectangle(500, 130, 430, 130));
		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(100.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(18.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, 1);

		slide.addShape(content);

		// =======content Platform========//
		Platform(slide, 2);
	}

	public void slide10hashtag(HSLFSlideShow ppt, HSLFSlide slide, ArrayList<String> rankhtg, int jml_htg)
			throws IOException {
		setBackground(slide, 2);

		HSLFTextBox title = new HSLFTextBox();

		title.setText("\n\n10 Hashtag\nYang Paling\nSering Digunakan");

		title.setAnchor(new java.awt.Rectangle(23, 30, 860, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.JUSTIFY);
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

		// ------------- add picture grafik -----------//
		HSLFPictureData pd8 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barcharthtgIGv1.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew8 = new HSLFPictureShape(pd8);
		pictNew8.setAnchor(new java.awt.Rectangle(430, 120, 490, 320));
		slide.addShape(pictNew8);

		// ---------content2---------//

		HSLFTextBox content2 = new HSLFTextBox();

		content2.appendText("Hashtag yang paling sering digunakan adalah ", false);
		content2.appendText(rankhtg.get(0) + " dengan jumlah " + String.valueOf(jml_htg) + " kali digunakan.", false);
		content2.setAnchor(new java.awt.Rectangle(23, 380, 380, 50));

		HSLFTextParagraph titleP2 = content2.getTextParagraphs().get(0);
		titleP2.setAlignment(TextAlign.JUSTIFY);
		titleP2.setSpaceAfter(0.);
		titleP2.setSpaceBefore(0.);
		titleP2.setLineSpacing(110.0);

		HSLFTextRun runTitle2 = titleP2.getTextRuns().get(0);
		runTitle2.setFontSize(22.);
		runTitle2.setFontFamily("Calibri");
		runTitle2.setBold(false);
		setTextColor(runTitle2, 2);

		slide.addShape(content2);

		// =======content Platform========//
		Platform(slide, 1);
	}

	public void slideNetworkHashtag3(HSLFSlide slide) {
		setBackground(slide, 1);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network\nHashtag");
		title.setAnchor(new java.awt.Rectangle(500, 10, 300, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.LEFT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(48.);
		runTitle.setFontFamily("Times New Roman");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, 1);

		slide.addShape(title);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();
		String jml_nodes = String.format("%s", "86");
		String jml_edges = String.format("%s", "78");
		String issue = String.format("%s", "PKI");
		String akun1 = String.format("%s", "@dan");
		String akun2 = String.format("%s", "@de");
		String akun3 = String.format("%s", "@tyu");
		String akun4 = String.format("%s", "@jrhe");
		String akun5 = String.format("%s", "@ewhg");
		String hashtag = String.format("%s", "54");
		String account = String.format("%s", "58");

		content.appendText("Lima akun yang paling sering menggunakan hashtag dalam issue : ", false);
		content.appendText(issue, false);
		content.appendText(" ini diantaranya: ", true);
		content.appendText("• " + akun1, true);
		content.appendText("• " + akun2, true);
		content.appendText("• " + akun3, true);
		content.appendText("• " + akun4, true);
		content.appendText("• " + akun5, true);

		content.appendText(" ", true);
		content.appendText("Jumlah Nodes	: ", true);
		content.appendText(jml_nodes, false);
		content.appendText("Jumlah Edges: ", true);
		content.appendText(jml_edges, false);
		content.appendText("Hashtag	: ", true);
		content.appendText(hashtag, false);
		content.appendText(" %", false);
		content.appendText("Account	: ", true);
		content.appendText(account, false);
		content.appendText(" %", false);

		content.setAnchor(new java.awt.Rectangle(500, 130, 430, 130));
		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(18.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, 1);

		slide.addShape(content);

		// =======content Platform========//
		Platform(slide, 2);
	}

	public void slideNetworkHashtag4(HSLFSlide slide) {
		setBackground(slide, 1);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network\nHashtag");
		title.setAnchor(new java.awt.Rectangle(500, 10, 300, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.LEFT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(48.);
		runTitle.setFontFamily("Times New Roman");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, 1);

		slide.addShape(title);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();
		String persentase_modularity_tertinggi_1 = String.format("%s", "50");
		String persentase_modularity_tertinggi_2 = String.format("%s", "40");
		String persentase_modularity_tertinggi_3 = String.format("%s", "30");
		String persentase_modularity_tertinggi_4 = String.format("%s", "20");
		String persentase_modularity_tertinggi_5 = String.format("%s", "10");
		String jml_nodes = String.format("%s", "86");
		String jml_edges = String.format("%s", "78");
		String account_1_kelompok_1 = String.format("%s", "@aku1");
		String account_2_kelompok_1 = String.format("%s", "@aku2");
		String account_3_kelompok_1 = String.format("%s", "@aku3");
		String account_1_kelompok_2 = String.format("%s", "@aku1");
		String account_2_kelompok_2 = String.format("%s", "@aku2");
		String account_3_kelompok_2 = String.format("%s", "@aku3");
		String account_1_kelompok_3 = String.format("%s", "@aku1");
		String account_2_kelompok_3 = String.format("%s", "@aku2");
		String account_3_kelompok_3 = String.format("%s", "@aku3");
		String account_1_kelompok_4 = String.format("%s", "@aku1");
		String account_2_kelompok_4 = String.format("%s", "@aku2");
		String account_3_kelompok_4 = String.format("%s", "@aku3");
		String account_1_kelompok_5 = String.format("%s", "@aku1");
		String account_2_kelompok_5 = String.format("%s", "@aku2");
		String account_3_kelompok_5 = String.format("%s", "@aku3");
		String issue = String.format("%s", "PKI");

		content.setText("Akun-akun yang paling sering menggunakan hashtag dalam issue ");
		content.appendText(issue, false);
		content.setText("Dari masing masing kelompok diantaranya: ");
		content.appendText("• Kelompok 1	: ", true);
		content.appendText(persentase_modularity_tertinggi_1 + " %", false);
		content.appendText(" ", true);
		content.appendText(account_1_kelompok_1 + ", " + account_2_kelompok_1 + ", " + account_3_kelompok_1, false);
		content.appendText("• Kelompok 2	: ", true);
		content.appendText(persentase_modularity_tertinggi_2 + " %", false);
		content.appendText(" ", true);
		content.appendText(account_1_kelompok_2 + ", " + account_2_kelompok_2 + ", " + account_3_kelompok_2, false);
		content.appendText("• Kelompok 3	: ", true);
		content.appendText(persentase_modularity_tertinggi_3 + " %", false);
		content.appendText(" ", true);
		content.appendText(account_1_kelompok_3 + ", " + account_2_kelompok_3 + ", " + account_3_kelompok_3, false);
		content.appendText("• Kelompok 4	: ", true);
		content.appendText(persentase_modularity_tertinggi_4 + " %", false);
		content.appendText(" ", true);
		content.appendText(account_1_kelompok_4 + ", " + account_2_kelompok_4 + ", " + account_3_kelompok_4, false);
		content.appendText("• Kelompok 5	: ", true);
		content.appendText(persentase_modularity_tertinggi_5 + " %", false);
		content.appendText(" ", true);
		content.appendText(account_1_kelompok_5 + ", " + account_2_kelompok_5 + ", " + account_3_kelompok_5, false);
		String hashtag = String.format("%s", "54");
		String account = String.format("%s", "58");

		content.appendText(" ", true);

		content.appendText("Jumlah Nodes	: ", true);
		content.appendText(jml_nodes, false);
		content.appendText("Jumlah Edges: ", true);
		content.appendText(jml_edges, false);
		content.appendText("Hashtag	: ", true);
		content.appendText(hashtag, false);
		content.appendText("Account	: ", true);
		content.appendText(account, false);

		content.setAnchor(new java.awt.Rectangle(500, 130, 420, 130));
		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(18.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, 1);

		slide.addShape(content);

		// =======content Platform========//
		Platform(slide, 1);
	}

	public void slideNetworkbyMention1(HSLFSlide slide) {
		setBackground(slide, 1);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network by\nMention");
		title.setAnchor(new java.awt.Rectangle(500, 10, 300, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.LEFT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(48.);
		runTitle.setFontFamily("Times New Roman");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, 1);

		slide.addShape(title);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();
		String jml_nodes = String.format("%s", "86");
		String jml_edges = String.format("%s", "78");
		String issue = String.format("%s", "PKI");
		String akun1 = String.format("%s", "@dan");
		String akun2 = String.format("%s", "@de");
		String akun3 = String.format("%s", "@tyu");
		String akun4 = String.format("%s", "@jrhe");
		String akun5 = String.format("%s", "@ewhg");

		content.appendText(
				"Network by mention merupakan network yang terdiri dari akun-akun yang di mention dalam issue ", false);
		content.appendText(issue + ".", false);
		content.appendText("Akun yang paling banyak di mention dalam issue ini adalah :", true);
		content.appendText("\n\t• " + akun1 + "\n\t• " + akun2 + "\n\t• " + akun3 + "\n\t• " + akun4 + "\n\t• " + akun5,
				false);

		content.appendText(" ", true);
		content.appendText("\nJumlah Nodes\t: ", true);
		content.appendText(jml_nodes, false);
		content.appendText("Jumlah Edges\t: ", true);
		content.appendText(jml_edges, false);

		content.setAnchor(new java.awt.Rectangle(500, 130, 420, 130));

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(18.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, 1);

		slide.addShape(content);

		// =======content Platform========//
		Platform(slide, 1);
	}

	public void slideNetworkbyMention2(HSLFSlide slide) {
		setBackground(slide, 1);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network by\nMention");

		title.setAnchor(new java.awt.Rectangle(23, 30, 860, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.LEFT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(48.);
		runTitle.setFontFamily("Times New Roman");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, 1);

		slide.addShape(title);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();
		String jml_nodes = String.format("%s", "86");
		String jml_edges = String.format("%s", "78");
		String issue = String.format("%s", "PKI");
		String akun1 = String.format("%s", "@dan");
		String akun2 = String.format("%s", "@de");
		String akun3 = String.format("%s", "@tyu");
		String akun4 = String.format("%s", "@jrhe");
		String akun5 = String.format("%s", "@ewhg");

		content.appendText("Akun-akun yang sering me-mention terkait ", false);
		content.appendText(issue + " :", true);
		content.appendText("  • " + akun1, true);
		content.appendText("  • " + akun2, true);
		content.appendText("  • " + akun3, true);
		content.appendText("  • " + akun4, true);
		content.appendText("  • " + akun5, true);

		content.appendText(" ", true);
		content.appendText("Jumlah Nodes\t: ", true);
		content.appendText(jml_nodes, false);
		content.appendText("Jumlah Edges\t: ", true);
		content.appendText(jml_edges, false);

		content.setAnchor(new java.awt.Rectangle(10, 170, 350, 385));
		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(18.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, 1);

		slide.addShape(content);

		// =======content Platform========//
		Platform(slide, 1);
	}

	public void slideKesimpulan(HSLFSlide slide, String[] time1, String[] time2) {
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
		runTitle.setItalic(true);
		setTextColor(runTitle, 2);

		slide.addShape(title);

		// -------------------------Content1------------------------ //

		HSLFTextBox content1 = new HSLFTextBox();

		content1.setText("Berdasarkan data sampel yang disediakan oleh Instagram pada tanggal " + time1[2] + " "
				+ time1[1] + " " + time1[5] + " hingga " + time2[2] + " " + time2[1] + " " + time2[5]
				+ ", diakses pada " + time2[2] + " " + time2[1] + " " + time2[5] + " dapat disimpulkan bahwa:");
		content1.setAnchor(new java.awt.Rectangle(20, 85, 920, 400));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.JUSTIFY);
		contentP1.setSpaceAfter(0.);
		contentP1.setSpaceBefore(0.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(18.);
		runContent1.setFontFamily("Corbel");
		runContent1.setBold(false);
		setTextColor(runContent1, 2);

		slide.addShape(content1);

		// =======content Platform========//
		Platform(slide, 1);
	}

}
