package com.ebdesk.report.slide;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
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

public class SNATwitterReport {
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

	public SNATwitterReport(String filename, int version, int width, int height, String tema, String db,
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

		String Title = "Amin Rais";
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

		float jml_tweets = list.size();
		float retweet = 0;
		float tweet = 0;
		float reply = 0;
		float quoted = 0;

		HashSet<String> set = new HashSet<String>();

		// ---------------time-----------------//
		ArrayList<String> settanggal = new ArrayList<String>(); // tanggal
		Map<String, Integer> map = new HashMap<String, Integer>();

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

		ArrayList<String> listhtg = new ArrayList<String>();
		List<String> myString = new ArrayList<String>();

		for (int ii = 0; ii < list.size(); ii++) {

			String g8 = (String) list.get(ii).get("text"); // text
			Pattern rex = Pattern.compile("(?<!\\S)\\p{Alpha}+(?!\\S)"); // words
			try {
				Matcher mx = rex.matcher(g8);
				while (mx.find()) {
					myString.add(mx.group(0));
				}
			} catch (Exception e) {

			}

			ArrayList x7 = (ArrayList) list.get(ii).get("hashtag"); // hashtag
			try {
				if (x7.size() != 0) {
					for (int ii1 = 0; ii1 < x7.size(); ii1++) {
						listhtg.add("#" + String.valueOf(x7.get(ii1)));

					}
				}
			} catch (Exception e) {
			}

			ArrayList y7 = (ArrayList) list.get(ii).get("hahstag"); // hashtag
			try {
				if (y7.size() != 0) {
					for (int ii1 = 0; ii1 < y7.size(); ii1++) {
						listhtg.add("#" + String.valueOf(y7.get(ii1)));
					}
				}
			} catch (Exception e) {
			}

			// ======time=========//
			g1 = (String) list.get(ii).get("created_at").toString();
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

			// ==========type data========//
			if (String.valueOf(list.get(ii).get("type")).equals("retweet")) {
				retweet += 1;
			} else if (String.valueOf(list.get(ii).get("type")).equals("quote")) {
				quoted += 1;
			} else if (String.valueOf(list.get(ii).get("type")).equals("reply")) {
				reply += 1;
			} else if (String.valueOf(list.get(ii).get("type")).equals("tweet")) {
				tweet += 1;
			}

			// ===================ID================//
			if (list.get(ii).get("user_id") != null) {
				set.add(String.valueOf(list.get(ii).get("user_id")));
			}
			if (list.get(ii).get("retweeted_user_id") != null) {
				set.add(String.valueOf(list.get(ii).get("retweeted_user_id")));
			}
			if (list.get(ii).get("quoted_user_id") != null) {
				set.add(String.valueOf(list.get(ii).get("quoted_user_id")));
			}

		}

		// =========tanggal========//

		String ga = (String) list.get(index1).get("created_at").toString();
		time1 = ga.split("\\s"); // tanggal awal
		String gk = (String) list.get(index2).get("created_at").toString();
		time2 = gk.split("\\s"); // tanggal akhir

		// ==================jumlah akun & Grafik=================//
		int jml_akun = set.size();
		float persen_retweet = retweet * 100 / jml_tweets;
		float persen_quoted = quoted * 100 / jml_tweets;
		float persen_reply = reply * 100 / jml_tweets;
		float persen_tweet = tweet * 100 / jml_tweets;
		float persentase_tertinggi = persen_tweet;
		String jenis_tweet_tertinggi = "Tweet";
		// y
		if (persentase_tertinggi < persen_retweet) {
			persentase_tertinggi = persen_retweet;
			jenis_tweet_tertinggi = "Retweet";
		}
		if (persentase_tertinggi < persen_quoted) {
			jenis_tweet_tertinggi = "Quoted";
			persentase_tertinggi = persen_quoted;
		}
		if (persentase_tertinggi < persen_reply) {
			jenis_tweet_tertinggi = "Reply";
			persentase_tertinggi = persen_reply;
		}

		PieChartReport chart1 = new PieChartReport();
		chart1.ringchart("Tweet", persen_tweet, "Reply", persen_reply, "Quoted", persen_quoted, "Retweet",
				persen_retweet, "pieTWv1.png", Color.white);

		// ================Time Series==========//

		String[] time3;
		final TimeSeries series = new TimeSeries("Random Data");
		int indexmap;
		HashSet<String> settime1 = new HashSet<String>(settanggal);
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
		JFreeChart timechart = ChartFactory.createTimeSeriesChart("Trend Tweet for Keyword: " + Title, "Time",
				"Tweet Count", dataset, false, false, false);
		timechart.setBackgroundPaint(Color.white);
		XYPlot plot = (XYPlot) timechart.getPlot();
		plot.setBackgroundPaint(Color.white);
		File timeChart = new File("TimeChartTWv1.png");
		ChartUtilities.saveChartAsJPEG(timeChart, timechart, 580, 370);
		int rata_rata_tweets = (int) (jml_tweets / (settime1.size() * 24));

		// ==========barchart======//
		BarChart chart7 = new BarChart();
		ArrayList<String> ranktgl = chart7.bar(settanggal, settime1, null, null, 1, 0, null);

		int jml_tweet_puncak = Collections.frequency(settanggal, String.valueOf(ranktgl.get(0))) / 24;
		String tanggal_puncak = String.valueOf(ranktgl.get(0));

		// =========wordcloud=========//
		List<String> myStringhtg = new ArrayList<String>(listhtg);
		Wordcloud word = new Wordcloud(); // wordcloud
		List<WordFrequency> wc = word.readCNN("wcTWv1.png", myString, "blue");
		List<WordFrequency> wchtg = word.readCNN("wchtgTWv1.png", myStringhtg, "blue");

		// ==================slide===============//
		HSLFSlideShow ppt = new HSLFSlideShow();
		ppt.setPageSize(new java.awt.Dimension(width, height));

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\twitter_logo.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(25, 503, 30, 30));

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
		slideTrend(ppt, slide3, time1, time2, list.size(), jml_akun, jml_tweet_puncak, tanggal_puncak, wchtg, wc);
		slide3.addShape(pictNew);

		// =======content2========//
		HSLFTextBox content2 = new HSLFTextBox();

		content2.setText(String.valueOf(jml_akun));
		content2.appendText(" accounts", false);
		content2.appendText(String.valueOf(list.size()), true);
		content2.appendText(" tweets\n", false);
		content2.appendText(String.valueOf(rata_rata_tweets), true);
		content2.appendText(" tweets/hour", false);

		content2.setAnchor(new java.awt.Rectangle(85, 290, 200, 385));

		HSLFTextParagraph titleP2 = content2.getTextParagraphs().get(0);
		titleP2.setAlignment(TextAlign.JUSTIFY);
		titleP2.setSpaceAfter(0.);
		titleP2.setSpaceBefore(0.);
		titleP2.setLineSpacing(105.0);

		HSLFTextRun runTitle2 = titleP2.getTextRuns().get(0);
		runTitle2.setFontSize(13.);
		runTitle2.setFontFamily("Calibri");
		runTitle2.setBold(true);
		setTextColor(runTitle2, version);
		runTitle2.setFontColor(Color.red);

		slide3.addShape(content2);

		// Slide 4
		HSLFSlide slide4 = ppt.createSlide();
		slideTweetTypes(ppt, slide4, persen_retweet, persen_quoted, persen_reply, persen_tweet, persentase_tertinggi,
				jenis_tweet_tertinggi);
		slide4.addShape(pictNew);

		// Slide 5
		HSLFSlide slide5 = ppt.createSlide();
		slideNetworkRetweet(slide5);
		slide5.addShape(pictNew);

		// // ------------- add picturegephi -----------//
		// HSLFPictureData pd3g = ppt.addPicture(new
		// File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\gephi.png"),
		// PictureData.PictureType.PNG);
		// HSLFPictureShape pictNew3g = new HSLFPictureShape(pd3g);
		// pictNew3g.setAnchor(new java.awt.Rectangle(40, 70, 460, 430));
		// slide5.addShape(pictNew3g);

		// Slide 6
		HSLFSlide slide6 = ppt.createSlide();
		slideNetworkRetweett(slide6);
		slide6.addShape(pictNew);

		// // ------------- add picturegephi -----------//
		// HSLFPictureData pd3gx = ppt.addPicture(new
		// File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\gephi.png"),
		// PictureData.PictureType.PNG);
		// HSLFPictureShape pictNew3gx = new HSLFPictureShape(pd3gx);
		// pictNew3gx.setAnchor(new java.awt.Rectangle(40, 70, 460, 430));
		// slide6.addShape(pictNew3gx);

		// Slide 7
		HSLFSlide slide7 = ppt.createSlide();
		slideNetworkMention(slide7);
		slide7.addShape(pictNew);

		// // ------------- add picturegephi -----------//
		// HSLFPictureData pd3gx1 = ppt.addPicture(new
		// File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\gephi2.png"),
		// PictureData.PictureType.PNG);
		// HSLFPictureShape pictNew3gx1 = new HSLFPictureShape(pd3gx1);
		// pictNew3gx1.setAnchor(new java.awt.Rectangle(45, 50, 450, 430));
		// slide7.addShape(pictNew3gx1);

		// Slide 8
		HSLFSlide slide8 = ppt.createSlide();
		slideConclusion(slide8, time1, time2);
		slide8.addShape(pictNew);

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

	public void Platform(HSLFSlide slide) {

		HSLFTextBox contentp = new HSLFTextBox();

		contentp.setText("Platform: Twitter");

		contentp.setAnchor(new java.awt.Rectangle(50, 510, 350, 50));

		HSLFTextParagraph titlePp = contentp.getTextParagraphs().get(0);
		titlePp.setAlignment(TextAlign.JUSTIFY);
		titlePp.setSpaceAfter(0.);
		titlePp.setSpaceBefore(0.);
		titlePp.setLineSpacing(110.0);

		HSLFTextRun runTitlep = titlePp.getTextRuns().get(0);
		runTitlep.setFontSize(15.);
		runTitlep.setFontFamily("Times New Roman");
		runTitlep.setBold(false);
		runTitlep.setItalic(true);
		runTitlep.setFontColor(new Color(61, 140, 255));

		slide.addShape(contentp);
	}

	public void slideMaster(HSLFSlide slide, String[] time2) {
		setBackground(slide, version);

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
		setTextColor(runTitle1, version);
		slide.addShape(title);

		// --------------------Investigation--------------------
		HSLFTextBox issue = new HSLFTextBox();
		issue.setText("Issue Investigation : " + tema);
		issue.setAnchor(new java.awt.Rectangle(50, 350, 900, 50));
		HSLFTextParagraph issueP = issue.getTextParagraphs().get(0);
		HSLFTextRun runIssue = issueP.getTextRuns().get(0);
		runIssue.setFontSize(32.);
		runIssue.setFontFamily("Times New Roman");
		runIssue.setBold(true);
		runIssue.setItalic(true);
		setTextColor(runIssue, version);
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
		setTextColor(runTitlept, version);

		slide.addShape(contentpt);

		// =======content Platform========//
		Platform(slide);
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
		content.appendText("' sedang ramai diperbincangkan di media social twitter.", false);
		content.setAnchor(new java.awt.Rectangle(280, 165, 400, 120));
		content.setLineColor(Color.BLUE);

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
		Platform(slide);
	}

	public void slideTrend(HSLFSlideShow ppt, HSLFSlide slide, String[] time1, String[] time2, int len_json,
			int jml_akun, int jml_tweet_puncak, String tanggal_puncak, List<WordFrequency> wchtg,
			List<WordFrequency> wc) throws Exception {
		// ---------judul---------//
		setBackground(slide, 2);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Trend");
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
		// ------------- add picture grafik -----------//
		HSLFPictureData pd3 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\TimeChartTWv1.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3 = new HSLFPictureShape(pd3);
		pictNew3.setAnchor(new java.awt.Rectangle(30, 270, 480, 220));
		slide.addShape(pictNew3);

		// ------------- add picture hashtag -----------//
		HSLFPictureData pd1 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\wchtgTWv1.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
		pictNew1.setAnchor(new java.awt.Rectangle(70, 60, 180, 180));
		slide.addShape(pictNew1);

		// ------------- add picture wordcloud -----------//
		HSLFPictureData pd2 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\wcTWv1.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew2 = new HSLFPictureShape(pd2);
		pictNew2.setAnchor(new java.awt.Rectangle(300, 60, 180, 180));
		slide.addShape(pictNew2);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();

		content.setText("Puncak pembicaraan tertinggi dari issue ");
		content.appendText(tema, false);
		content.appendText(" terjadi pada tanggal ", false);
		content.appendText(tanggal_puncak, false);
		content.appendText(" dengan jumlah pembicaraan mencapai ", false);
		content.appendText(String.valueOf(jml_tweet_puncak), false);
		content.appendText(" tweets/hours.", false);

		content.appendText("\nJumlah pembicaraan issue ini dalam kurun ", false);
		content.appendText(time1[2] + "/" + time1[1] + "/" + time1[5], false);
		content.appendText(" - ", false);
		content.appendText(time2[2] + "/" + time2[1] + "/" + time2[5], false);
		content.appendText(" yakni sebanyak ", false);
		content.appendText(String.valueOf(len_json), false);
		content.appendText(" tweets, dengan jumlah massa twitter yang membicarakan sebanyak ", false);
		content.appendText(String.valueOf(jml_akun), false);
		content.appendText(" akun.", false);

		content.appendText(" ", true);
		content.appendText("Keywords yang paling banyak digunakan adalah ", true);
		content.appendText(wc.get(0).getWord(), false);
		content.appendText(", ", false);
		content.appendText(wc.get(1).getWord(), false);
		content.appendText(", ", false);
		content.appendText(wc.get(2).getWord(), false);
		content.appendText(", ", false);
		content.appendText(wc.get(3).getWord(), false);
		content.appendText(" dan ", false);
		content.appendText(wc.get(4).getWord(), false);
		content.appendText(". Hashtag yang paling sering digunakan adalah hashtag " + wchtg.get(0).getWord() + ", ",
				false);
		content.appendText(wchtg.get(1).getWord() + " dan " + wchtg.get(2).getWord() + ".", false);
		content.setAnchor(new java.awt.Rectangle(530, 90, 400, 340));
		content.setLineColor(Color.BLUE);

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(18.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, 2);

		slide.addShape(content);
		// =======keterangan============//

		HSLFTextBox contentp1 = new HSLFTextBox();

		contentp1.setText("Ket: besar tulisan menunjukan tingkat frekuensi kata/hashtag tersebut di tweet.");

		contentp1.setAnchor(new java.awt.Rectangle(530, 450, 400, 70));

		HSLFTextParagraph titlePp1 = contentp1.getTextParagraphs().get(0);
		titlePp1.setAlignment(TextAlign.JUSTIFY);
		titlePp1.setSpaceAfter(0.);
		titlePp1.setSpaceBefore(0.);
		titlePp1.setLineSpacing(110.0);

		HSLFTextRun runTitlep1 = titlePp1.getTextRuns().get(0);
		runTitlep1.setFontSize(14.);
		runTitlep1.setFontFamily("Calibri");
		runTitlep1.setBold(false);
		setTextColor(runTitlep1, 2);

		slide.addShape(contentp1);

		// =======content Platform========//
		Platform(slide);

		// =======content hashtag========//
		HSLFTextBox contentp2 = new HSLFTextBox();

		contentp2.setText("Hashtag");

		contentp2.setAnchor(new java.awt.Rectangle(130, 240, 200, 50));

		HSLFTextParagraph titlePp2 = contentp2.getTextParagraphs().get(0);
		titlePp2.setAlignment(TextAlign.JUSTIFY);
		titlePp2.setSpaceAfter(0.);
		titlePp2.setSpaceBefore(0.);
		titlePp2.setLineSpacing(110.0);

		HSLFTextRun runTitlep2 = titlePp2.getTextRuns().get(0);
		runTitlep2.setFontSize(15.);
		runTitlep2.setFontFamily("Calibri");
		runTitlep2.setBold(false);
		setTextColor(runTitlep2, 2);

		slide.addShape(contentp2);

		// =======content Wordcloud========//
		HSLFTextBox contentp21 = new HSLFTextBox();

		contentp21.setText("Wordcloud");

		contentp21.setAnchor(new java.awt.Rectangle(350, 240, 200, 50));

		HSLFTextParagraph titlePp21 = contentp21.getTextParagraphs().get(0);
		titlePp21.setAlignment(TextAlign.JUSTIFY);
		titlePp21.setSpaceAfter(0.);
		titlePp21.setSpaceBefore(0.);
		titlePp21.setLineSpacing(110.0);

		HSLFTextRun runTitlep21 = titlePp21.getTextRuns().get(0);
		runTitlep21.setFontSize(15.);
		runTitlep21.setFontFamily("Calibri");
		runTitlep21.setBold(false);
		setTextColor(runTitlep21, 2);

		slide.addShape(contentp21);

	}

	public void slideTweetTypes(HSLFSlideShow ppt, HSLFSlide slide, float persen_retweet, float persen_quoted,
			float persen_reply, float persen_tweet, float persentase_tertinggi, String jenis_tweet_tertinggi)
			throws IOException {
		// -----------------judul----------------//
		setBackground(slide, 2);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Tweet Types");
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

		// ----- add picture-------------//
		HSLFPictureData pd4 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\pieTWv1.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew4 = new HSLFPictureShape(pd4);
		pictNew4.setAnchor(new java.awt.Rectangle(260, 110, 400, 310));
		slide.addShape(pictNew4);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();

		content.setText("Hasil diagram tersebut menunjukkan bahwa mayoritas penyebaran issue dilakukan melalui ");
		content.appendText(jenis_tweet_tertinggi, false);
		content.appendText(" dengan persentase sebesar ", false);
		content.appendText(String.valueOf(persentase_tertinggi), false);
		content.appendText(" %.", false);

		content.setAnchor(new java.awt.Rectangle(280, 450, 400, 70));
		content.setLineColor(Color.BLUE);

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(16.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, 2);

		slide.addShape(content);

		// =======content Platform========//
		Platform(slide);

	}

	public void slideNetworkRetweet(HSLFSlide slide) {
		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network of Retweet");
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
		setTextColor(runTitle, version);

		slide.addShape(title);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();

		String jml_kelompok_diatas3persen = String.format("%s", "568");
		String persentase_modularity_tertinggi_1 = String.format("%s", "50");
		String persentase_modularity_tertinggi_2 = String.format("%s", "40");
		String persentase_modularity_tertinggi_3 = String.format("%s", "30");
		String persentase_modularity_tertinggi_4 = String.format("%s", "20");
		String persentase_modularity_tertinggi_lain = String.format("%s", "10");
		String jml_nodes = String.format("%s", "86");
		String jml_retweet = String.format("%s", "78");

		content.setText(
				"Network di samping menunjukkan kelompok berdasarkan aktivitas retweet yang menyatakan respon dari netizen di media sosial twitter.");
		content.appendText(" ", true);
		content.appendText("Hasil visualisasi di samping menunjukkan bahwa terdapat ", true);
		content.appendText(jml_kelompok_diatas3persen, false);
		content.appendText(
				" kelompok besar dalam jaringan, dengan besar persentase tiap kelompok dalam jaringan sebagai berikut:",
				false);

		content.appendText(" ", true);
		content.appendText("• Kelompok 1	: ", true);
		content.appendText(persentase_modularity_tertinggi_1, false);
		content.appendText("• Kelompok 2	: ", true);
		content.appendText(persentase_modularity_tertinggi_2, false);
		content.appendText("• Kelompok 3	: ", true);
		content.appendText(persentase_modularity_tertinggi_3, false);
		content.appendText("• Kelompok 4	: ", true);
		content.appendText(persentase_modularity_tertinggi_4, false);
		content.appendText("• Kelompok Lain	: ", true);
		content.appendText(persentase_modularity_tertinggi_lain, false);
		content.appendText(" ", true);
		content.appendText(" ", true);
		content.appendText("Jumlah Akun 	: ", true);
		content.appendText(jml_nodes, false);
		content.appendText("Jumlah Retweet 	: ", true);
		content.appendText(jml_retweet, false);

		content.setAnchor(new java.awt.Rectangle(530, 90, 400, 385));
		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(15.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, version);

		slide.addShape(content);

		// =======content Platform========//
		Platform(slide);
	}

	public void slideNetworkRetweett(HSLFSlide slide) {
		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network of Retweet");
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
		setTextColor(runTitle, version);

		slide.addShape(title);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();

		String akun_in_degree_tertinggi_1_di_kelompok_1 = String.format("%s", "surya");
		String akun_in_degree_tertinggi_2_di_kelompok_1 = String.format("%s", "malik");
		String akun_in_degree_tertinggi_3_di_kelompok_1 = String.format("%s", "reza");
		String akun_in_degree_tertinggi_1_di_kelompok_2 = String.format("%s", "bayu");
		String akun_in_degree_tertinggi_2_di_kelompok_2 = String.format("%s", "ashif");
		String akun_in_degree_tertinggi_3_di_kelompok_2 = String.format("%s", "dicky");
		String akun_in_degree_tertinggi_1_di_kelompok_3 = String.format("%s", "aris");
		String akun_in_degree_tertinggi_2_di_kelompok_3 = String.format("%s", "ulil");
		String akun_in_degree_tertinggi_3_di_kelompok_3 = String.format("%s", "tomo");
		String akun_in_degree_tertinggi_1_di_kelompok_4 = String.format("%s", "iqbal");
		String akun_in_degree_tertinggi_2_di_kelompok_4 = String.format("%s", "samsu");
		String akun_in_degree_tertinggi_3_di_kelompok_4 = String.format("%s", "duban");
		String akun_in_degree_tertinggi_1_di_kelompok_l = String.format("%s", "hendra");
		String akun_in_degree_tertinggi_2_di_kelompok_l = String.format("%s", "sani");
		String akun_in_degree_tertinggi_3_di_kelompok_l = String.format("%s", "damas");
		String jml_akun = String.format("%s", "86");
		String jml_tweets = String.format("%s", "78");

		content.setText(
				"Hasil visualisasi di samping menunjukkan bahwa terdapat beberapa akun influencer di setiap kelompok pada jaringan, yakni:");
		content.appendText(" ", true);

		content.appendText(" ", true);
		content.appendText("• Kelompok 1	: ", true);
		content.appendText(akun_in_degree_tertinggi_1_di_kelompok_1, false);
		content.appendText(akun_in_degree_tertinggi_2_di_kelompok_1, false);
		content.appendText(akun_in_degree_tertinggi_3_di_kelompok_1, false);
		content.appendText("• Kelompok 2	: ", true);
		content.appendText(akun_in_degree_tertinggi_1_di_kelompok_2, false);
		content.appendText(akun_in_degree_tertinggi_2_di_kelompok_2, false);
		content.appendText(akun_in_degree_tertinggi_3_di_kelompok_2, false);
		content.appendText("• Kelompok 3	: ", true);
		content.appendText(akun_in_degree_tertinggi_1_di_kelompok_3, false);
		content.appendText(akun_in_degree_tertinggi_2_di_kelompok_3, false);
		content.appendText(akun_in_degree_tertinggi_3_di_kelompok_3, false);
		content.appendText("• Kelompok 4	: ", true);
		content.appendText(akun_in_degree_tertinggi_1_di_kelompok_4, false);
		content.appendText(akun_in_degree_tertinggi_2_di_kelompok_4, false);
		content.appendText(akun_in_degree_tertinggi_3_di_kelompok_4, false);
		content.appendText("• Kelompok Lain	: ", true);
		content.appendText(akun_in_degree_tertinggi_1_di_kelompok_l, false);
		content.appendText(akun_in_degree_tertinggi_2_di_kelompok_l, false);
		content.appendText(akun_in_degree_tertinggi_3_di_kelompok_l, false);
		content.appendText(" ", true);
		content.appendText(" ", true);
		content.appendText("Jumlah Akun 	: ", true);
		content.appendText(jml_akun, false);
		content.appendText("Jumlah Retweet 	: ", true);
		content.appendText(jml_tweets, false);

		content.setAnchor(new java.awt.Rectangle(530, 90, 400, 385));
		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(15.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, version);

		slide.addShape(content);

		// =======content Platform========//
		Platform(slide);
	}

	public void slideNetworkMention(HSLFSlide slide) {
		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network of Mention");
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
		setTextColor(runTitle, version);

		slide.addShape(title);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();

		String nama_isu = String.format("%s", "Malik Burhanuddin");
		String akun1 = String.format("%s", "Jokowi");
		String akun2 = String.format("%s", "detikcom");
		String akun3 = String.format("%s", "faizasalassegaf");
		String akun4 = String.format("%s", "CNNIndonesi");
		String akun5 = String.format("%s", "VIVACoid");
		String mention1 = String.format("%s", "50");
		String mention2 = String.format("%s", "40");
		String mention3 = String.format("%s", "60");
		String mention4 = String.format("%s", "70");
		String mention5 = String.format("%s", "80");
		String jml_akun = String.format("%s", "86");
		String jml_mention = String.format("%s", "78");

		content.setText(
				"Dengan menggunakan Equal (edges_kind): Mention dan Giant Component, maka akan terlihat akun-akun yang memiliki frekuensi tertinggi disebut dalam tweet.");
		content.appendText(" ", true);
		content.appendText("Akun yang paling sering disebut terkait issue ", true);
		content.appendText(nama_isu, false);
		content.appendText(" adalah:", false);

		content.appendText("• " + akun1 + "\t\t: " + mention1 + " Mention", true);
		content.appendText("• " + akun2 + "\t\t: " + mention2 + " Mention", true);
		content.appendText("• " + akun3 + "\t\t: " + mention3 + " Mention", true);
		content.appendText("• " + akun4 + "\t\t: " + mention4 + " Mention", true);
		content.appendText("• " + akun5 + "\t\t: " + mention5 + " Mention", true);
		content.appendText("\n\nJumlah Akun 	: " + jml_akun + "\nJumlah Mention 	: " + jml_mention, true);

		content.setAnchor(new java.awt.Rectangle(530, 90, 400, 385));
		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(15.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, version);

		slide.addShape(content);

		// =======content Platform========//
		Platform(slide);

		// =======content keterangan========//
		HSLFTextBox contentpk = new HSLFTextBox();

		contentpk.setText("Warna diurutkan berdasarkan frekuensi mention yang diterima oleh suatu akun.");

		contentpk.setAnchor(new java.awt.Rectangle(50, 478, 500, 50));

		HSLFTextParagraph titlePpk = contentpk.getTextParagraphs().get(0);
		titlePpk.setAlignment(TextAlign.JUSTIFY);
		titlePpk.setSpaceAfter(0.);
		titlePpk.setSpaceBefore(0.);
		titlePpk.setLineSpacing(110.0);

		HSLFTextRun runTitlepk = titlePpk.getTextRuns().get(0);
		runTitlepk.setFontSize(13.);
		runTitlepk.setFontFamily("Times New Roman");
		runTitlepk.setBold(false);
		setTextColor(runTitlepk, version);

		slide.addShape(contentpk);

	}

	public void slideConclusion(HSLFSlide slide, String[] time1, String[] time2) {
		setBackground(slide, 2);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Conclusion");
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

		// =======content Platform========//
		Platform(slide);

		// -------------------------Content1------------------------ //

		HSLFTextBox content1 = new HSLFTextBox();

		content1.setText("Berdasarkan data sampel yang disediakan oleh Twitter pada tanggal " + time1[2] + " "
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
		setTextColor(runContent1, version);

		slide.addShape(content1);

	}

}
