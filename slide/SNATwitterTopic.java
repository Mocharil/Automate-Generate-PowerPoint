package com.ebdesk.report.slide;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
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
import java.util.Map;
import java.util.Set;

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

public class SNATwitterTopic {

	private int version;
	private String filename;
	private int width;
	private int height;
	private String tema;
	private static String mongoHost = "192.168.114.162";
	private static String mongoPort = "27017";
	private static String mongoUsername = "dev";
	private static String mongoPassword = "Rahas!adev20!8";
	private static String mongoSource = "admin";

	public SNATwitterTopic(String filename, int version, int width, int height, String tema) {
		this.filename = filename;
		this.version = version;
		this.height = height;
		this.width = width;
		this.tema = tema;
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

		MongoDatabase database = mongoClient.getDatabase("criteria_twitter");
		MongoCollection<Document> collection = database.getCollection("00da342174b57cd6ec28e6bfceb93be1");

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
		int retweet = 0;
		int tweet = 0;
		int reply = 0;
		int quoted = 0;

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

		// Document doc=(Document) list.get(0).get("hashtag");
		// System.out.println(doc);
		ArrayList<String> listhtg = new ArrayList<String>();

		HashSet<String> setID = new HashSet<String>();
		Map<String, Integer> mapID = new HashMap<String, Integer>();

		for (int ii = 0; ii < list.size(); ii++) {
			ArrayList x7 = (ArrayList) list.get(ii).get("hashtag"); // hashtag
			ArrayList y7 = (ArrayList) list.get(ii).get("hahstag"); // hashtag

			try {
				if (x7.size() != 0) {
					for (int ii1 = 0; ii1 < x7.size(); ii1++) {
						listhtg.add(String.valueOf(x7.get(ii1))); // hashtag

					}
				}
			} catch (Exception e) {
			}

			try {
				if (y7.size() != 0) {
					for (int ii1 = 0; ii1 < y7.size(); ii1++) {
						listhtg.add(String.valueOf(y7.get(ii1))); // hashtag
					}
				}
			} catch (Exception e) {
			}

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

			if (String.valueOf(list.get(ii).get("type")).equals("retweet")) {
				retweet += 1;
			}

			else if (String.valueOf(list.get(ii).get("type")).equals("quote")) {
				quoted += 1;
			}

			else if (String.valueOf(list.get(ii).get("type")).equals("reply")) {
				reply += 1;
			}

			else if (String.valueOf(list.get(ii).get("type")).equals("tweet")) {
				tweet += 1;
			}

			if (list.get(ii).get("user_id") != null) {
				set.add(String.valueOf(list.get(ii).get("user_id")));
				setID.add(String.valueOf(list.get(ii).get("user_id")));
				mapID.put(String.valueOf(list.get(ii).get("user_id")), ii);

			}
			if (list.get(ii).get("retweeted_user_id") != null) {
				set.add(String.valueOf(list.get(ii).get("retweeted_user_id")));
			}
			if (list.get(ii).get("quoted_user_id") != null) {
				set.add(String.valueOf(list.get(ii).get("quoted_user_id")));
			}

		}

		HashSet<String> settime1 = new HashSet<String>(settanggal);
		ArrayList<String> ranktgl = new ArrayList<String>();

		String ga = (String) list.get(index1).get("created_at").toString();
		time1 = ga.split("\\s");
		String gk = (String) list.get(index2).get("created_at").toString();
		time2 = gk.split("\\s");

		// ==================jumlah akun================= //
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

		// -----------grafik------------//
		grafik chart = new grafik();
		chart.mulai(null, "Tweet", persen_tweet, "Reply", persen_reply, "Quoted", persen_quoted, "Retweet",
				persen_retweet, "pieTW.png", 4, "black");

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

		ArrayList<String> listindex = new ArrayList<String>();
		ArrayList<Integer> listfollower = new ArrayList<Integer>();

		Map<String, Integer> mapfollower = new HashMap<String, Integer>();
		Map<String, Integer> mapreach = new HashMap<String, Integer>();
		Map<String, Integer> maptweet = new HashMap<String, Integer>();

		ArrayList<String> listindexfollower = new ArrayList<String>();
		ArrayList<String> listindextweet = new ArrayList<String>();
		ArrayList<String> listindexreach = new ArrayList<String>();

		Object[] px = setID.toArray();
		for (int f = 0; f < setID.size(); f++) {
			int tw = (Integer) list.get(Integer.valueOf(mapID.get(String.valueOf(px[f])))).get("user_statuses_count");
			int fol = (Integer) list.get(Integer.valueOf(mapID.get(String.valueOf(px[f])))).get("user_followers_count");
			listindex.add(String.valueOf(mapID.get(String.valueOf(px[f])))); // Index
			mapfollower.put(String.valueOf(mapID.get(String.valueOf(px[f]))),
					(Integer) list.get(Integer.valueOf(mapID.get(String.valueOf(px[f])))).get("user_followers_count")); // index,
																														// followers
			maptweet.put(String.valueOf(mapID.get(String.valueOf(px[f]))),
					(Integer) list.get(Integer.valueOf(mapID.get(String.valueOf(px[f])))).get("user_statuses_count")); // index,
																														// tweet
			mapreach.put(String.valueOf(mapID.get(String.valueOf(px[f]))), Math.abs(tw * fol)); // index, reach

		}

		Map<String, Integer> mapsort = sortByValues(mapfollower); // sorted by value
		Set set2 = mapsort.entrySet();
		Iterator iterator2 = set2.iterator();
		while (iterator2.hasNext()) {
			Map.Entry me2 = (Map.Entry) iterator2.next();
			listindexfollower.add(String.valueOf(me2.getKey()));
		}

		Map<String, Integer> mapsorttweet = sortByValues(maptweet); // sorted by value
		Set set2tweet = mapsorttweet.entrySet();
		Iterator iterator2tweet = set2tweet.iterator();
		while (iterator2tweet.hasNext()) {
			Map.Entry me2tweet = (Map.Entry) iterator2tweet.next();
			listindextweet.add(String.valueOf(me2tweet.getKey()));
		}

		Map<String, Integer> mapsortreach = sortByValues(mapreach); // sorted by value
		Set set2reach = mapsortreach.entrySet();
		Iterator iterator2reach = set2reach.iterator();
		while (iterator2reach.hasNext()) {
			Map.Entry me2reach = (Map.Entry) iterator2reach.next();
			listindexreach.add(String.valueOf(me2reach.getKey()));
		}

		table_(list, listindexfollower, "tabelfollower.png");
		table_(list, listindextweet, "tabeltweet.png");
		table_(list, listindexreach, "tabelreach.png");

		final XYDataset dataset = (XYDataset) new TimeSeriesCollection(series);
		JFreeChart timechart = ChartFactory.createTimeSeriesChart("Trend Tweet for Keyword: " + tema, "Time",
				"Tweet Count", dataset, false, false, false);
		timechart.setBackgroundPaint(Color.white);
		XYPlot plot = (XYPlot) timechart.getPlot();
		plot.setBackgroundPaint(new Color(255, 255, 255, 0));

		HashSet<String> sethtg = new HashSet<String>(listhtg);

		File timeChart = new File("TimeChartTW.png");
		ChartUtilities.saveChartAsJPEG(timeChart, timechart, 580, 370);
		int rata_rata_tweets = (int) (jml_tweets / (settime1.size() * 24));

		ArrayList<String> rankhtg = new ArrayList<String>();

		BarChart chart7 = new BarChart();
		ranktgl = chart7.bar(settanggal, settime1, null, null, 1, 0, null);
		int jml_tweet_puncak = Collections.frequency(settanggal, String.valueOf(ranktgl.get(0))) / 24;
		String tanggal_puncak = String.valueOf(ranktgl.get(0));
		rankhtg = chart7.bar(listhtg, sethtg, null, null, 10, 0, null);

		// ==================slide===============//
		HSLFSlideShow ppt = new HSLFSlideShow();
		ppt.setPageSize(new java.awt.Dimension(width, height));

		// Slide 1
		HSLFSlide slide1 = ppt.createSlide();
		slideMaster(slide1, time2);

		// Slide 2
		HSLFSlide slide2 = ppt.createSlide();
		slideInformation(slide2, time1, time2);

		// Slide 3
		HSLFSlide slide3 = ppt.createSlide();
		slidekeywords(slide3, time2);

		// Slide 4
		HSLFSlide slide4 = ppt.createSlide();
		slideTrend(slide4, time1, time2, jml_tweet_puncak, tanggal_puncak);

		// ------------- add picture grafik -----------//
		HSLFPictureData pd3 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\TimeChartTW.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3 = new HSLFPictureShape(pd3);
		pictNew3.setAnchor(new java.awt.Rectangle(30, 30, 670, 260));
		slide4.addShape(pictNew3);

		// =======content2========//
		HSLFTextBox content2 = new HSLFTextBox();
		content2.appendText("Data. Sum\t: " + String.valueOf(list.size()) + " tweets\n", false);
		content2.appendText("Averange\t: " + String.valueOf(rata_rata_tweets) + " tweets/hour\n", false);
		content2.appendText("Acc. Sum\t: " + String.valueOf(jml_akun) + " accounts", false);

		content2.setAnchor(new java.awt.Rectangle(120, 60, 200, 120));

		HSLFTextParagraph titleP2 = content2.getTextParagraphs().get(0);
		titleP2.setAlignment(TextAlign.JUSTIFY);
		titleP2.setSpaceAfter(0.);
		titleP2.setSpaceBefore(0.);
		titleP2.setLineSpacing(105.0);

		HSLFTextRun runTitle2 = titleP2.getTextRuns().get(0);
		runTitle2.setFontSize(13.);
		runTitle2.setFontFamily("Calibri");
		runTitle2.setBold(true);
		runTitle2.setFontColor(Color.red);

		slide4.addShape(content2);

		// Slide 5
		HSLFSlide slide5 = ppt.createSlide();
		slideTweetTypes(slide5, persentase_tertinggi, jenis_tweet_tertinggi, time1, time2, list.size(), retweet, tweet,
				quoted, reply, jml_akun);

		// ----- add picture-------------//
		HSLFPictureData pd4 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\pieTW.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew4 = new HSLFPictureShape(pd4);
		pictNew4.setAnchor(new java.awt.Rectangle(510, 120, 360, 275));
		slide5.addShape(pictNew4);

		// Slide 6

		HSLFSlide slide6 = ppt.createSlide();
		slideTopTweets(slide6, time2);

		// Slide 7

		HSLFSlide slide7 = ppt.createSlide();
		slideTopAccount(slide7, time2);
		// ------------- add picture -----------//
		HSLFPictureData pd3f = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\tabelfollower.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3f = new HSLFPictureShape(pd3f);
		pictNew3f.setAnchor(new java.awt.Rectangle(250, 30, 220, 190));
		slide7.addShape(pictNew3f);

		// ------------- add picture -----------//
		HSLFPictureData pd3t = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\tabeltweet.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3t = new HSLFPictureShape(pd3t);
		pictNew3t.setAnchor(new java.awt.Rectangle(490, 30, 220, 190));
		slide7.addShape(pictNew3t);

		// ------------- add picture -----------//
		HSLFPictureData pd3r = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\tabelreach.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3r = new HSLFPictureShape(pd3r);
		pictNew3r.setAnchor(new java.awt.Rectangle(730, 30, 220, 190));
		slide7.addShape(pictNew3r);

		// Slide 8
		HSLFSlide slide8 = ppt.createSlide();
		slideConclusion(slide8);

		FileOutputStream out = new FileOutputStream(filename);
		ppt.write(out);
		out.close();
		System.out.println("-------------------> Done <-------------------");
	}

	private static HashMap sortByValues(Map<String, Integer> mapfollower) {
		List list = new LinkedList(mapfollower.entrySet());
		// Defined Custom Comparator here
		Collections.sort(list, new Comparator() {
			public int compare(Object o1, Object o2) {
				return ((Comparable) ((Map.Entry) (o1)).getValue()).compareTo(((Map.Entry) (o2)).getValue());
			}
		});

		// Here I am copying the sorted list in HashMap
		// using LinkedHashMap to preserve the insertion order
		HashMap sortedHashMap = new LinkedHashMap();
		for (Iterator it = list.iterator(); it.hasNext();) {
			Map.Entry entry = (Map.Entry) it.next();
			sortedHashMap.put(entry.getKey(), entry.getValue());
		}
		return sortedHashMap;
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

	private void table_(ArrayList<Document> list, ArrayList<String> listindexfollower, String namefile) {
		int k = 0;
		String[][] data = new String[10][4];
		String headers[] = { "Count", "Follower", "Username", "Reach" };
		for (int v = listindexfollower.size() - 1; v > listindexfollower.size() - 11; v--) { // bawah
			int ko = 0;
			k += 1;
			for (int c = 0; c < 4; c++) {
				try {
					int tweet = (Integer) list.get(Integer.valueOf(listindexfollower.get(v)))
							.get("user_statuses_count");
					int followers = (Integer) list.get(Integer.valueOf(listindexfollower.get(v)))
							.get("user_followers_count");

					if (ko == 0) {
						data[k - 1][ko] = String.valueOf(tweet);
					} else if (ko == 1) {
						data[k - 1][ko] = String.valueOf(followers);
					} else if (ko == 2) {
						data[k - 1][ko] = String
								.valueOf(list.get(Integer.valueOf(listindexfollower.get(v))).get("username"));
					} else if (ko == 3) {
						data[k - 1][ko] = String.valueOf(Math.abs(tweet * followers));
					}
					ko += 1;
				} catch (Exception e) {
				}
			}
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
		setTextColor(runTitlep, version);

		slide.addShape(contentp);
	}

	public void slideMaster(HSLFSlide slide, String[] time2) {
		HSLFShape shape = new HSLFAutoShape(ShapeType.RECT);
		shape.setAnchor(new java.awt.Rectangle(0, 0, width, height));
		HSLFFill fill = shape.getFill();
		fill.setFillType(HSLFFill.FILL_SHADE);
		fill.setBackgroundColor(Color.cyan);
		fill.setForegroundColor(Color.black);
		slide.addShape(shape);

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
		issue.setText(tema);
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
		setBackground(slide, version);

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
		setTextColor(runTitle, version);

		slide.addShape(title);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();

		content.setText("Dalam kurun " + time1[2] + "/" + time1[1] + "/" + time1[5]);
		content.appendText(" - ", false);
		content.appendText(time2[2] + "/" + time2[1] + "/" + time2[5], false);
		content.appendText(", issue mengenai '", false);
		content.appendText(tema, false);
		content.appendText("' sedang ramai diperbincangkan di media social twitter.", false);
		content.setAnchor(new java.awt.Rectangle(280, 165, 400, 130));
		content.setLineColor(Color.YELLOW);

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(22.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, version);

		slide.addShape(content);

		// =======content Platform========//
		Platform(slide);
	}

	public void slidekeywords(HSLFSlide slide, String[] time2) throws Exception {
		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Keywords");
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

		String key1 = String.format("%s", "HTILanjutkanPerjuangan");
		String key2 = String.format("%s", "7MeiHTIMenang");
		String key3 = String.format("%s", "IslamSelamatkanNegeri");
		String key4 = String.format("%s", "UmatBersamaHTI");
		String key5 = String.format("%s", "HTILayakMenang");

		HSLFTextBox content = new HSLFTextBox();

		content.setText("Keyword yang digunakan untuk topic " + tema);

		content.setAnchor(new java.awt.Rectangle(310, 120, 340, 60));
		content.setLineColor(Color.YELLOW);

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.CENTER);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(20.);
		runContent.setFontFamily("Corbel");
		runContent.setBold(true);
		setTextColor(runContent, version);

		slide.addShape(content);

		// ========key=========//
		HSLFTextBox content1 = new HSLFTextBox();

		content1.appendText(key1 + "\n" + key2 + "\n" + key3 + "\n" + key4 + "\n" + key5, false);

		content1.setAnchor(new java.awt.Rectangle(310, 180, 340, 140));
		content1.setLineColor(Color.YELLOW);

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.CENTER);
		contentP1.setSpaceAfter(0.);
		contentP1.setSpaceBefore(0.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(18.);
		runContent1.setFontFamily("Corbel");
		runContent1.setBold(false);
		setTextColor(runContent1, version);

		slide.addShape(content1);

		// =========ket======//

		HSLFTextBox contentpk = new HSLFTextBox();

		contentpk.appendText("Data yang digunakan berdasarkan data yang disediakan oleh twitter, diakses pada tanggal\n"
				+ time2[2] + " " + time2[1] + " " + time2[5], false);

		contentpk.setAnchor(new java.awt.Rectangle(130, 430, 700, 50));
		contentpk.setLineColor(Color.YELLOW);
		HSLFTextParagraph titlePpk = contentpk.getTextParagraphs().get(0);
		titlePpk.setAlignment(TextAlign.CENTER);
		titlePpk.setSpaceAfter(0.);
		titlePpk.setSpaceBefore(0.);
		titlePpk.setLineSpacing(110.0);

		HSLFTextRun runTitlepk = titlePpk.getTextRuns().get(0);
		runTitlepk.setFontSize(18.);
		runTitlepk.setFontFamily("Calibri");
		runTitlepk.setBold(false);
		setTextColor(runTitlepk, version);

		slide.addShape(contentpk);

		// =======content Platform========//
		Platform(slide);

	}

	public void slideTrend(HSLFSlide slide, String[] time1, String[] time2, int jml_tweet_puncak, String tanggal_puncak)
			throws Exception {
		setBackground(slide, version);

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
		setTextColor(runTitle, version);

		slide.addShape(title);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();

		content.setText("Dalam kurun waktu ");
		content.appendText(time1[2] + "/" + time1[1] + "/" + time1[5] + " - ", false);
		content.appendText(time2[2] + "/" + time2[1] + "/" + time2[5] + ", puncak pembicaraan terkait topik " + tema,
				false);
		content.appendText(" mencapai puncak trend pada " + tanggal_puncak + ".", false);
		content.setAnchor(new java.awt.Rectangle(20, 295, 920, 220));
		content.setLineColor(Color.YELLOW);

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(18.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, version);

		slide.addShape(content);

		// =======content Platform========//
		Platform(slide);

	}

	public void slideTweetTypes(HSLFSlide slide, float persentase_tertinggi, String jenis_tweet_tertinggi,
			String[] time1, String[] time2, int jml_tweet, int retweet, int tweet, int quoted, int reply, int jml_akun)
			throws IOException {

		setBackground(slide, version);

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
		setTextColor(runTitle, version);

		slide.addShape(title);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();

		content.setText("Dari data tweet ");
		content.appendText(time1[2] + "/" + time1[1] + "/" + time1[5] + " hingga ", false);
		content.appendText(
				time2[2] + "/" + time2[1] + "/" + time2[5] + ", secara keseluruhan, tweet di dominasi oleh tipe ",
				false);
		content.appendText(jenis_tweet_tertinggi, false);
		content.appendText(" dengan persentase sebesar ", false);
		content.appendText(String.valueOf(persentase_tertinggi) + " %.", false);
		content.appendText(
				"\n\n Lebih banyak netizen yang tertarik untuk merespon tweet dibandingkan membuat tweet terkait topik ini.",
				false);

		content.setAnchor(new java.awt.Rectangle(50, 160, 400, 180));
		content.setLineColor(Color.YELLOW);

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(18.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, version);

		slide.addShape(content);

		// =======content2========//
		HSLFTextBox content2 = new HSLFTextBox();
		content2.appendText("Data. Sum \t: " + String.valueOf(jml_tweet) + " tweets\n", false);
		content2.appendText("     Retweet \t: " + String.valueOf(retweet), false);
		content2.appendText("\n     Tweet \t: " + String.valueOf(tweet), false);
		content2.appendText("\n     Quoted \t: " + String.valueOf(quoted), false);
		content2.appendText("\n     Reply \t: " + String.valueOf(reply), false);
		content2.appendText("\nAcc. Sum \t: " + String.valueOf(jml_akun) + "accounts", false);

		content2.setAnchor(new java.awt.Rectangle(140, 350, 300, 300));

		HSLFTextParagraph titleP2 = content2.getTextParagraphs().get(0);
		titleP2.setAlignment(TextAlign.JUSTIFY);
		titleP2.setSpaceAfter(0.);
		titleP2.setSpaceBefore(0.);
		titleP2.setLineSpacing(105.0);

		HSLFTextRun runTitle2 = titleP2.getTextRuns().get(0);
		runTitle2.setFontSize(16.);
		runTitle2.setFontFamily("Calibri");
		runTitle2.setBold(true);
		setTextColor(runTitle2, version);

		slide.addShape(content2);

		// ===ket===//

		HSLFTextBox contentp1 = new HSLFTextBox();

		contentp1.setText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + time2[2] + "/"
				+ time2[1] + "/" + time2[5]);

		contentp1.setAnchor(new java.awt.Rectangle(50, 480, 600, 50));

		HSLFTextParagraph titlePp1 = contentp1.getTextParagraphs().get(0);
		titlePp1.setAlignment(TextAlign.JUSTIFY);
		titlePp1.setSpaceAfter(0.);
		titlePp1.setSpaceBefore(0.);
		titlePp1.setLineSpacing(110.0);

		HSLFTextRun runTitlep1 = titlePp1.getTextRuns().get(0);
		runTitlep1.setFontSize(14.);
		runTitlep1.setFontFamily("Calibri");
		runTitlep1.setBold(false);
		setTextColor(runTitlep1, version);

		slide.addShape(contentp1);

		// =======content Platform========//
		Platform(slide);

	}

	public void slideTopTweets(HSLFSlide slide, String[] time2) throws Exception {
		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Top Tweets");
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

		String key1 = String.format("%s", "HTILanjutkanPerjuangan");

		HSLFTextBox content = new HSLFTextBox();

		content.setText("5 Tweet terkait " + tema + " yang paling banyak di respon netizen");
		content.appendText("\nTweet didominasi oleh " + key1 + " dan terlihat tweet yang paling banyak di-retweet ",
				false);
		content.appendText("dan disukai adalah promosi tweet dari tagar" + tema + " itu sendiri", false);

		content.setAnchor(new java.awt.Rectangle(40, 360, 880, 80));
		content.setLineColor(Color.YELLOW);

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.CENTER);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(20.);
		runContent.setFontFamily("Corbel");
		runContent.setBold(false);
		setTextColor(runContent, version);

		slide.addShape(content);

		// ===ket===//

		HSLFTextBox contentp1 = new HSLFTextBox();

		contentp1.setText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + time2[2] + "/"
				+ time2[1] + "/" + time2[5]);

		contentp1.setAnchor(new java.awt.Rectangle(50, 480, 530, 50));

		HSLFTextParagraph titlePp1 = contentp1.getTextParagraphs().get(0);
		titlePp1.setAlignment(TextAlign.JUSTIFY);
		titlePp1.setSpaceAfter(0.);
		titlePp1.setSpaceBefore(0.);
		titlePp1.setLineSpacing(110.0);

		HSLFTextRun runTitlep1 = titlePp1.getTextRuns().get(0);
		runTitlep1.setFontSize(14.);
		runTitlep1.setFontFamily("Calibri");
		runTitlep1.setBold(false);
		setTextColor(runTitlep1, version);

		slide.addShape(contentp1);

		// =======content Platform========//
		Platform(slide);

	}

	public void slideTopAccount(HSLFSlide slide, String[] time2) throws Exception {
		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Top Account");
		title.setAnchor(new java.awt.Rectangle(50, 30, 200, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.LEFT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(40.);
		runTitle.setFontFamily("Times New Roman");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);

		// -------------------------Content------------------------ //

		String key1 = String.format("%s", "HTILanjutkanPerjuangan");

		HSLFTextBox content = new HSLFTextBox();

		content.setText("Dalam topik ini,");
		content.appendText(
				"\n• akun yang berkontribusi dalam topik ini dengan jumlah follower terbanyak adalah " + key1, false);
		content.appendText(
				" artinya akun tersebut diasumsikan memiliki masa yang banyak untuk melihat tweet yang dibuatnya.",
				false);
		content.appendText(
				"\n• akun yang berkontribusi dalam topik ini dengan membuat tweet paling banyak adalah " + key1, false);
		content.appendText(
				"\n• Akun yang berkontribusi dalam topik ini dengan nilai reach (tweet x follower) paling besar adalah"
						+ key1,
				false);
		content.appendText(
				" dapat diasumsikan bahwa jika " + key1 + " membuat tweet, tweet tersebut akan cepat tersebar.", false);

		content.setAnchor(new java.awt.Rectangle(30, 300, 880, 155));
		content.setLineColor(Color.YELLOW);

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(17.);
		runContent.setFontFamily("Corbel");
		runContent.setBold(false);
		setTextColor(runContent, version);

		slide.addShape(content);

		// ===caption==//

		HSLFTextBox contentc = new HSLFTextBox();

		contentc.setText("Data 10 akun paling popular.\nDilihat dari banyaknya follower.");

		contentc.setAnchor(new java.awt.Rectangle(250, 230, 220, 190));

		HSLFTextParagraph contentPc = contentc.getTextParagraphs().get(0);
		contentPc.setAlignment(TextAlign.JUSTIFY);
		contentPc.setSpaceAfter(0.);
		contentPc.setSpaceBefore(0.);
		contentPc.setLineSpacing(110.0);

		HSLFTextRun runContentc = contentPc.getTextRuns().get(0);
		runContentc.setFontSize(16.);
		runContentc.setFontFamily("Corbel");
		runContentc.setBold(false);
		setTextColor(runContentc, version);

		slide.addShape(contentc);

		// ===caption==//

		HSLFTextBox contentcd = new HSLFTextBox();

		contentcd.setText("Data 10 akun paling banyak membuat tweet.");

		contentcd.setAnchor(new java.awt.Rectangle(490, 230, 220, 190));

		HSLFTextParagraph contentPcd = contentcd.getTextParagraphs().get(0);
		contentPcd.setAlignment(TextAlign.JUSTIFY);
		contentPcd.setSpaceAfter(0.);
		contentPcd.setSpaceBefore(0.);
		contentPcd.setLineSpacing(110.0);

		HSLFTextRun runContentcd = contentPcd.getTextRuns().get(0);
		runContentcd.setFontSize(16.);
		runContentcd.setFontFamily("Corbel");
		runContentcd.setBold(false);
		setTextColor(runContentcd, version);

		slide.addShape(contentcd);

		// ===caption==//

		HSLFTextBox contentce = new HSLFTextBox();

		contentce.setText("Data 10 berpengaruh dilihat dari reach (tweet x follower).");

		contentce.setAnchor(new java.awt.Rectangle(730, 230, 220, 190));

		HSLFTextParagraph contentPce = contentce.getTextParagraphs().get(0);
		contentPce.setAlignment(TextAlign.JUSTIFY);
		contentPce.setSpaceAfter(0.);
		contentPce.setSpaceBefore(0.);
		contentPce.setLineSpacing(110.0);

		HSLFTextRun runContentce = contentPce.getTextRuns().get(0);
		runContentce.setFontSize(16.);
		runContentce.setFontFamily("Corbel");
		runContentce.setBold(false);
		setTextColor(runContentce, version);

		slide.addShape(contentce);
		// ===ket===//

		HSLFTextBox contentp1 = new HSLFTextBox();

		contentp1.setText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + time2[2] + "/"
				+ time2[1] + "/" + time2[5]);

		contentp1.setAnchor(new java.awt.Rectangle(45, 470, 530, 50));

		HSLFTextParagraph titlePp1 = contentp1.getTextParagraphs().get(0);
		titlePp1.setAlignment(TextAlign.JUSTIFY);
		titlePp1.setSpaceAfter(0.);
		titlePp1.setSpaceBefore(0.);
		titlePp1.setLineSpacing(110.0);

		HSLFTextRun runTitlep1 = titlePp1.getTextRuns().get(0);
		runTitlep1.setFontSize(14.);
		runTitlep1.setFontFamily("Calibri");
		runTitlep1.setBold(false);
		setTextColor(runTitlep1, version);

		slide.addShape(contentp1);

		// ======================ket2=================//
		HSLFTextBox contentp1z = new HSLFTextBox();

		contentp1z.setText("Ket: Impact merupakan hasil perkalian dari jumlah tweet tentang topik dengan follower.");
		contentp1z.setAnchor(new java.awt.Rectangle(45, 490, 530, 50));

		HSLFTextParagraph titlePp1z = contentp1z.getTextParagraphs().get(0);
		titlePp1z.setAlignment(TextAlign.JUSTIFY);
		titlePp1z.setSpaceAfter(0.);
		titlePp1z.setSpaceBefore(0.);
		titlePp1z.setLineSpacing(110.0);

		HSLFTextRun runTitlep1z = titlePp1z.getTextRuns().get(0);
		runTitlep1z.setFontSize(14.);
		runTitlep1z.setFontFamily("Calibri");
		runTitlep1z.setBold(false);
		setTextColor(runTitlep1z, version);

		slide.addShape(contentp1z);

		// =======content Platform========//
		Platform(slide);

	}

	public void slideConclusion(HSLFSlide slide) {
		setBackground(slide, version);

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
		setTextColor(runTitle, version);

		slide.addShape(title);

		// =======content Platform========//
		Platform(slide);
	}

}
