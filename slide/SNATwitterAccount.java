package com.ebdesk.report.slide;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.util.Date;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
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
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.imageio.ImageIO;
import javax.swing.BorderFactory;
import javax.swing.JPanel;

import org.apache.poi.hslf.usermodel.HSLFAutoShape;
import org.apache.poi.hslf.usermodel.HSLFFill;
import org.apache.poi.hslf.usermodel.HSLFLine;
import org.apache.poi.hslf.usermodel.HSLFPictureData;
import org.apache.poi.hslf.usermodel.HSLFPictureShape;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFTable;
import org.apache.poi.hslf.usermodel.HSLFTableCell;
import org.apache.poi.hslf.usermodel.HSLFTextBox;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.sl.draw.DrawTableShape;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.sl.usermodel.VerticalAlignment;
import org.bson.Document;
import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.XYPlot;
import org.jfree.data.time.Day;
import org.jfree.data.time.TimeSeries;
import org.jfree.data.time.TimeSeriesCollection;
import org.jfree.data.xy.XYDataset;

import com.github.davidmoten.rtree.geometry.Line;
import com.github.davidmoten.rtree.geometry.Rectangle;
import com.kennycason.kumo.WordFrequency;
import com.mongodb.DB;
import com.mongodb.DBCollection;
import com.mongodb.DBCursor;
import com.mongodb.DBObject;
import com.mongodb.MongoClient;
import com.mongodb.MongoClientURI;
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
import java.text.SimpleDateFormat;

public class SNATwitterAccount {

	private int version;
	private String filename;
	private int width;
	private int height;
	private String tema;
	private static String db;
	private static String collect;

	public SNATwitterAccount(String filename, int version, int width, int height, String tema, String db,
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
		MongoClient mongoClient = new MongoClient(new MongoClientURI("mongodb://192.168.99.30:27017"));
		DB database = mongoClient.getDB(db);
		DBCollection collection = database.getCollection(collect);
		DBCursor cursor = collection.find();
		ArrayList<DBObject> list = new ArrayList<DBObject>();

		try {

			while (cursor.hasNext()) {
				list.add(cursor.next());

			}
		} catch (Exception e) {
			System.out.println("gagal");
		}

		List<String> myString = new ArrayList<String>();

		// ---------------time-----------------//
		ArrayList<String> settanggal = new ArrayList<String>(); // tanggal
		Map<String, Integer> map = new HashMap<String, Integer>();

		String g1;
		String time1;
		String time2;
		String[] bulan = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };

		int nilai_bulan = 1;
		int tanggal1 = 1;
		int tanggal_awal = 999999999;
		int index1 = 0;
		int index2 = 0;
		int tanggal_akhir = 0;
		ArrayList<String> listhtg = new ArrayList<String>();

		int retweet = 0;
		int tweet = 0;
		int reply = 0;
		int quoted = 0;

		HashSet<String> setID = new HashSet<String>(); // akun yang mentweet

		Map<String, Integer> mapID = new HashMap<String, Integer>();

		ArrayList<String> listmention = new ArrayList<String>();
		ArrayList<Integer> indexinfo = new ArrayList<Integer>();
		for (int ii = 0; ii < list.size(); ii++) {
			if (String.valueOf(list.get(ii).get("username")).equals("shamsiali2")) {
				indexinfo.add(ii);

				// ===============TEXT==============//
				Pattern rex = Pattern.compile("(?<!\\S)\\p{Alpha}+(?!\\S)"); // regex words
				Pattern re = Pattern.compile("(#\\w+)"); // regex hashtag

				try {
					Matcher mx1 = rex.matcher((String) list.get(ii).get("text"));
					while (mx1.find()) {
						myString.add(mx1.group(0)); // Wordcloud
					}

				} catch (Exception e) {

				}
				ArrayList<Object> tx = (ArrayList<Object>) list.get(ii).get("hashtag"); // hashtag
				if (tx.size() >= 1) {
					for (int c = 0; c < tx.size(); c++) {
						listhtg.add((String) tx.get(c));
					}
				}

				// ------------time---------//
				SimpleDateFormat dateFormat = new SimpleDateFormat("dd MMMM yyyy", Locale.ENGLISH);
				dateFormat.setTimeZone(TimeZone.getTimeZone("UTC"));
				time1 = dateFormat.format(list.get(ii).get("created_at"));

				tanggal1 = Integer.valueOf((String) list.get(ii).get("pub_day"));
				settanggal.add(time1);
				map.put(time1, ii);

				if (tanggal_awal > tanggal1) {
					tanggal_awal = tanggal1;
					index1 = ii;
				}

				if (tanggal_akhir < tanggal1) {
					tanggal_akhir = tanggal1;
					index2 = ii;
				}

				if (list.get(ii).get("retweeted_user_screen_name") == null) {

					// ============MENTION==============//
					ArrayList<Object> x = (ArrayList<Object>) list.get(ii).get("mention");
					if (x.size() >= 1) {
						for (int c = 0; c < x.size(); c++) {
							DBObject h = (DBObject) x.get(c);
							listmention.add("@" + h.get("username"));
						}
					}
					setID.add(String.valueOf(list.get(ii).get("tweet_id")));
					mapID.put(String.valueOf(list.get(ii).get("tweet_id")), ii);

				}

				// =========type data========//

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

			}
		}

		// ==============tabel=============//
		Map<String, Integer> maplikes = new HashMap<String, Integer>();
		Map<String, Integer> mapretweet = new HashMap<String, Integer>();
		Map<String, Integer> mapreply = new HashMap<String, Integer>();

		Object[] px = setID.toArray();
		for (int f = 0; f < setID.size(); f++) {
			if (list.get(Integer.valueOf(mapID.get(String.valueOf(px[f])))).get("likes_count") != null) {
				maplikes.put(String.valueOf(mapID.get(String.valueOf(px[f]))),
						(Integer) list.get(Integer.valueOf(mapID.get(String.valueOf(px[f])))).get("likes_count"));

			}
			if (list.get(Integer.valueOf(mapID.get(String.valueOf(px[f])))).get("shares_count") != null) {
				mapretweet.put(String.valueOf(mapID.get(String.valueOf(px[f]))),
						(Integer) list.get(Integer.valueOf(mapID.get(String.valueOf(px[f])))).get("shares_count"));
			}
			if (list.get(Integer.valueOf(mapID.get(String.valueOf(px[f])))).get("comments_count") != null) {
				mapreply.put(String.valueOf(mapID.get(String.valueOf(px[f]))),
						(Integer) list.get(Integer.valueOf(mapID.get(String.valueOf(px[f])))).get("comments_count"));
			}
		}

		ArrayList<String> listindexlikes = new ArrayList<String>(sortbyValue(maplikes));
		ArrayList<String> listindexretweet = new ArrayList<String>(sortbyValue(mapretweet));
		ArrayList<String> listindexreply = new ArrayList<String>(sortbyValue(mapreply));

		HSLFTable table1 = table_(list, listindexlikes, "likes_count", "Likes Count");
		HSLFTable table2 = table_(list, listindexretweet, "shares_count", "Shares Count");
		HSLFTable table3 = table_(list, listindexreply, "comments_count", "Reply Count");

		// ==================Jumlah Akun & Grafik================= //
		float jml_tweets = retweet + quoted + reply + tweet;
		int jml_tweet = retweet + quoted + reply + tweet;
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
				persen_retweet, "donatTWacoount.png", Color.white);

		// ===============TANGGAL=============//
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd MMMM yyyy", Locale.ENGLISH);
		dateFormat.setTimeZone(TimeZone.getTimeZone("UTC"));
		time1 = dateFormat.format(list.get(index1).get("created_at"));
		time2 = dateFormat.format(list.get(index2).get("created_at"));

		// =============Time Series Chart===========//
		String[] time3;
		final TimeSeries series = new TimeSeries("Random Data");
		int indexmap;
		HashSet<String> settime1 = new HashSet<String>(settanggal);
		Object[] r = settime1.toArray();
		for (int oo = 0; oo < settime1.size(); oo++) {
			try {
				indexmap = map.get(String.valueOf(r[oo]));

				SimpleDateFormat dateFormat2 = new SimpleDateFormat("dd MMM yyyy", Locale.ENGLISH);
				dateFormat2.setTimeZone(TimeZone.getTimeZone("UTC"));
				g1 = dateFormat2.format(list.get(indexmap).get("created_at"));
				time3 = g1.split("\\s");
				for (int bul = 0; bul < 12; bul++) {

					if (String.valueOf(time3[1]).equals(String.valueOf(bulan[bul]))) {

						nilai_bulan = (bul + 1);
					}
				}

				series.add(new Day(Integer.valueOf(time3[0]), nilai_bulan, Integer.valueOf(time3[2])),
						new Double(Collections.frequency(settanggal, String.valueOf(r[oo]))));

			} catch (Exception e) {
			}
		}

		final XYDataset dataset = (XYDataset) new TimeSeriesCollection(series);
		JFreeChart timechart = ChartFactory.createTimeSeriesChart("Trend Tweet for Account: " + tema, "Time",
				"Tweet Count", dataset, false, false, false);
		timechart.setBackgroundPaint(Color.white);
		XYPlot plot = (XYPlot) timechart.getPlot();
		plot.setBackgroundPaint(new Color(255, 255, 255));
		File timeChart = new File("TimeChartTWAccount.png");
		ChartUtilities.saveChartAsJPEG(timeChart, timechart, 580, 370);

		// ==============BarChart=============//

		BarChart chart7 = new BarChart();
		ArrayList<String> ranktgl = chart7.bar(settanggal, settime1, null, null, 1, 0, null);
		ArrayList<String> rankmention = chart7.bar(listmention, new HashSet<String>(listmention),
				"barmentionTWaccount.png", "TOP Mention Account", 5, 1, "green");

		// ==============wordcloud============//

		List<String> myStringhtg = new ArrayList<String>(listhtg);
		System.out.println("mulai wordcloud");
		Wordcloud word = new Wordcloud();
		List<WordFrequency> wc = word.readCNN("wcTWaccount.png", myString, "blue");
		List<WordFrequency> wchtg = word.readCNN("wchtgTWaccount.png", myStringhtg, "blue");

		HSLFTable table4 = table2x10(wc);
		HSLFTable table5 = table2x10(wchtg);

		// ========================INFORMATION========================//

		try {
			saveImage(String.valueOf(list.get(0).get("user_profile_image_url")).replace("normal", "400x400"),
					"photoTWakun.png");
		} catch (Exception e) {
			saveImage(String.valueOf(list.get(0).get("user_profile_image_url")), "photoTWakun.png");
		}

		String nama_akun = (String) list.get(0).get("_criteria_id");

		String tanggal_puncak = String.valueOf(ranktgl.get(0));
		int jml_tweet_puncak = Collections.frequency(settanggal, String.valueOf(ranktgl.get(0)));
		int banyak_mention = Collections.frequency(listmention, rankmention.get(0));

		SimpleDateFormat dateFormat1 = new SimpleDateFormat("dd MM yyyy", Locale.ENGLISH);
		dateFormat1.setTimeZone(TimeZone.getTimeZone("UTC"));
		SimpleDateFormat myFormat = new SimpleDateFormat("dd MM yyyy");

		Date date1 = myFormat.parse(dateFormat1.format(list.get(index1).get("created_at")));
		Date date2 = myFormat.parse(dateFormat1.format(list.get(index2).get("created_at")));
		long diff = date2.getTime() - date1.getTime();
		long timeframe = TimeUnit.DAYS.convert(diff, TimeUnit.MILLISECONDS);

		// =============================SLIDE============================//

		HSLFSlideShow ppt = new HSLFSlideShow();
		ppt.setPageSize(new java.awt.Dimension(width, height));

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\twitter_logo.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(300, 276, 60, 60));

		// Slide 1
		HSLFSlide slide1 = ppt.createSlide();
		slideMaster(slide1, nama_akun);
		slide1.addShape(pictNew);

		// Slide 2
		HSLFSlide slide2 = ppt.createSlide();
		slideInformation(ppt, slide2, list, indexinfo.get(0));

		// Slide 3
		HSLFSlide slide4 = ppt.createSlide();
		slideTrendTweet(ppt, slide4, time1, time2, tanggal_puncak, timeframe, jml_tweet);

		// Slide 4
		HSLFSlide slide5 = ppt.createSlide();
		slideHashtagCloud(ppt, slide5, wchtg, time2);
		slide5.addShape(table5);
		table5.moveTo(30, 100);

		// Slide 5
		HSLFSlide slide3 = ppt.createSlide();
		slideWordcloud(ppt, slide3, wc, time2);
		slide3.addShape(table4);
		table4.moveTo(30, 100);

		// Slide 6
		HSLFSlide slide6 = ppt.createSlide();
		slideTweetTypes(slide6, time2, jenis_tweet_tertinggi, persentase_tertinggi);

		// ------------- add picture -----------//
		HSLFPictureData pd3q = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\donatTWacoount.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3q = new HSLFPictureShape(pd3q);
		pictNew3q.setAnchor(new java.awt.Rectangle(500, 130, 400, 310));
		slide6.addShape(pictNew3q);

		// Slide 7
		HSLFSlide slide7 = ppt.createSlide();
		slideTopPostbyRetweet(slide7, time2, nama_akun);
		slide7.addShape(table2);
		table2.moveTo(20, 100);

		// Slide 8
		HSLFSlide slide8 = ppt.createSlide();
		slideTopPostbyFavorite(slide8, time2, nama_akun);
		slide8.addShape(table1);
		table1.moveTo(20, 100);

		// Slide 9
		HSLFSlide slide9 = ppt.createSlide();
		slideTopMentionAccount(ppt, slide9, rankmention, time2, nama_akun, banyak_mention);

		// Slide 13
		HSLFSlide slide13 = ppt.createSlide();
		slideTopPostbyReply(slide13, time2, nama_akun);
		slide13.addShape(table3);
		table3.moveTo(20, 100);

		// Slide 9
		HSLFSlide slide91 = ppt.createSlide();
		slideOptimalTimeDiagram(slide91);

		// // ------------- add picture -----------//
		// HSLFPictureData pd14r = ppt.addPicture(
		// new
		// File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barHashtagTeleChannel.png"),
		// PictureData.PictureType.PNG);
		// HSLFPictureShape pictNew14r = new HSLFPictureShape(pd14r);
		// pictNew14r.setAnchor(new java.awt.Rectangle(500, 100, 400, 370));
		// slide91.addShape(pictNew14r);

		// Slide 10
		HSLFSlide slide22 = ppt.createSlide();
		slideConclusion(slide22, time1, time2);

		FileOutputStream out = new FileOutputStream(filename);
		ppt.write(out);
		out.close();
		System.out.println("-------------------> Done <-------------------");
	}

	private HSLFTable table_(ArrayList<DBObject> list, ArrayList<String> listindexlikes, String pp, String hh) {

		String headers[] = { "Date", hh, "Text", "Tweet Url" };
		String[][] data = new String[6][4];

		int k = 0;

		for (int v = listindexlikes.size() - 1; v > listindexlikes.size() - 7; v--) { // bawah
			int ko = 0;
			k += 1;
			for (int c = 0; c < 4; c++) {
				try {

					if (ko == 0) {

						String g8 = String
								.valueOf(list.get(Integer.valueOf(listindexlikes.get(v))).get("created_at").toString());
						String[] time7 = g8.split("\\s");

						data[k - 1][ko] = time7[2] + " " + time7[1] + " " + time7[5];
					} else if (ko == 1) {
						data[k - 1][ko] = String.valueOf(list.get(Integer.valueOf(listindexlikes.get(v))).get(pp));
					} else if (ko == 2) {

						data[k - 1][ko] = String.valueOf(list.get(Integer.valueOf(listindexlikes.get(v))).get("text"))
								.replaceAll("\n", " ");

					} else if (ko == 3) {

						data[k - 1][ko] = String.valueOf(list.get(Integer.valueOf(listindexlikes.get(v))).get("url"));

					}
					ko += 1;
				} catch (Exception e) {
				}
			}
		}

		HSLFTable table1 = new HSLFTable(6, 4);
		for (int i = 0; i < data.length; i++) {
			for (int j = 0; j < headers.length; j++) {
				HSLFTableCell cell = table1.getCell(i, j);
				HSLFTextRun rt = cell.getTextParagraphs().get(0).getTextRuns().get(0);
				rt.setFontFamily("Agency FB");
				rt.setFontSize(15.);

				if (i == 0) {
					cell.getFill().setForegroundColor(new Color(255, 255, 255));
					rt.setBold(true);
					try {
						cell.setText(headers[j]);

					} catch (Exception e) {
					}
				} else if (i % 2 != 0) {
					cell.setFillColor(new Color(245, 245, 245));
					try {
						cell.setText(data[i - 1][j]);
					} catch (Exception e) {
					}
				} else {
					cell.setFillColor(new Color(255, 255, 255));
					try {
						cell.setText(data[i - 1][j]);
					} catch (Exception e) {
					}
				}
				cell.setVerticalAlignment(VerticalAlignment.MIDDLE);
				cell.setHorizontalCentered(true);
			}
		}

		// set table borders
		HSLFLine border = table1.createBorder();
		border.setLineColor(Color.green);
		table1.setAllBorders(border);

		table1.setColumnWidth(0, 95);
		table1.setColumnWidth(1, 95);
		table1.setColumnWidth(2, 600);
		table1.setColumnWidth(3, 120);

		return table1;

	}

	private HSLFTable table2x10(List<WordFrequency> wc) {

		String headers[] = { "Word", "Count" };
		String[][] data = new String[8][2];

		int k = 0;

		for (int v = 0; v < 8; v++) { // bawah
			int ko = 0;
			k += 1;
			for (int c = 0; c < 2; c++) {
				try {

					if (ko == 0) {

						data[k - 1][ko] = wc.get(v).getWord();
					} else {
						data[k - 1][ko] = String.valueOf(wc.get(v).getFrequency());
					}
					ko += 1;
				} catch (Exception e) {
				}
			}
		}

		HSLFTable table1 = new HSLFTable(8, 2);
		for (int i = 0; i < data.length; i++) {
			for (int j = 0; j < headers.length; j++) {
				HSLFTableCell cell = table1.getCell(i, j);
				HSLFTextRun rt = cell.getTextParagraphs().get(0).getTextRuns().get(0);
				rt.setFontFamily("Agency FB");
				rt.setFontSize(16.);

				if (i == 0) {
					cell.getFill().setForegroundColor(new Color(255, 255, 255));
					rt.setBold(true);
					try {
						cell.setText(headers[j]);

					} catch (Exception e) {
					}
				} else if (i % 2 != 0) {
					cell.setFillColor(new Color(245, 245, 245));
					try {
						cell.setText(data[i - 1][j]);
					} catch (Exception e) {
					}
				} else {
					cell.setFillColor(new Color(255, 255, 255));
					try {
						cell.setText(data[i - 1][j]);
					} catch (Exception e) {
					}
				}
				cell.setVerticalAlignment(VerticalAlignment.MIDDLE);
				cell.setHorizontalCentered(true);
			}
		}

		// set table borders
		HSLFLine border = table1.createBorder();
		border.setLineColor(Color.green);
		table1.setAllBorders(border);

		table1.setColumnWidth(0, 170);
		table1.setColumnWidth(1, 80);

		return table1;

	}

	private static ArrayList<String> sortbyValue(Map<String, Integer> mapretweet) {
		ArrayList<String> listindexretweet = new ArrayList<String>();
		Map<String, Integer> mapsortretweet = sortByValues(mapretweet); // sorted by value
		Set set2tweet = mapsortretweet.entrySet();
		Iterator iterator2retweet = set2tweet.iterator();

		while (iterator2retweet.hasNext()) {
			Map.Entry me2retweet = (Map.Entry) iterator2retweet.next();
			listindexretweet.add(String.valueOf(me2retweet.getKey()));
		}

		return (listindexretweet);
	}

	private static HashMap sortByValues(Map<String, Integer> mapfollower) {
		List list = new LinkedList(mapfollower.entrySet());
		Collections.sort(list, new Comparator() {
			public int compare(Object o1, Object o2) {
				return ((Comparable) ((Map.Entry) (o1)).getValue()).compareTo(((Map.Entry) (o2)).getValue());
			}
		});

		HashMap sortedHashMap = new LinkedHashMap();
		for (Iterator it = list.iterator(); it.hasNext();) {
			Map.Entry entry = (Map.Entry) it.next();
			sortedHashMap.put(entry.getKey(), entry.getValue());
		}
		return sortedHashMap;
	}

	public static void saveImage(String imageUrl, String destinationFile) throws IOException {
		URL url = new URL(imageUrl);
		InputStream is = url.openStream();
		OutputStream os = new FileOutputStream(destinationFile);

		byte[] b = new byte[2048];
		int length;

		while ((length = is.read(b)) != -1) {
			os.write(b, 0, length);
		}

		is.close();
		os.close();
		System.out.println("selesai");
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
		title.setAnchor(new java.awt.Rectangle(10, 15, 940, 100));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);

		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);
		if (versi == 1) {
			titleP.setAlignment(TextAlign.RIGHT);
		} else if (versi == 7) {
			titleP.setAlignment(TextAlign.LEFT);
		} else {
			titleP.setAlignment(TextAlign.LEFT);
		}

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(54.);
		runTitle.setFontFamily("Agency FB");

		if (versi == 1) {
			runTitle.setFontColor(Color.BLACK);
		} else if (versi == 7) {
			runTitle.setFontColor(Color.BLACK);
		} else {
			runTitle.setFontColor(Color.WHITE);
		}

		slide.addShape(title);

	}

	public void slideMaster(HSLFSlide slide, String nama_akun) {
		setBackground(slide, 1);
		HSLFTextBox title = new HSLFTextBox();
		title.setText("SOCIAL");
		title.appendText("NETWORK", true);
		title.appendText("ANALYSIS", true);
		title.setAnchor(new java.awt.Rectangle(60, 120, 500, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.LEFT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);
		titleP.setLineSpacing(80.0);
		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(77.);
		runTitle.setFontFamily("Agency FB");

		runTitle.setFontColor(Color.red);

		HSLFTextParagraph titleP1 = title.getTextParagraphs().get(2);
		HSLFTextRun runTitle1 = titleP1.getTextRuns().get(0);
		runTitle1.setFontSize(77.);
		runTitle1.setFontFamily("Agency FB");
		runTitle1.setFontColor(Color.WHITE);
		slide.addShape(title);

		// --------------------Investigation--------------------
		HSLFTextBox issue = new HSLFTextBox();
		issue.appendText("INVESTIGASI Akun " + nama_akun, false);
		issue.setAnchor(new java.awt.Rectangle(60, 350, 900, 50));
		HSLFTextParagraph issueP = issue.getTextParagraphs().get(0);
		HSLFTextRun runIssue = issueP.getTextRuns().get(0);
		runIssue.setFontSize(24.);
		runIssue.setFontFamily("Century Schoolbook");
		runIssue.setFontColor(Color.WHITE);
		slide.addShape(issue);

		// -------------------------Date------------------------
		HSLFTextBox date = new HSLFTextBox();
		Date now = new Date(System.currentTimeMillis());
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
		date.setText(dateFormat.format(now));
		date.setAnchor(new java.awt.Rectangle(60, 470, 200, 50));
		HSLFTextParagraph dateP = date.getTextParagraphs().get(0);
		dateP.setAlignment(TextAlign.LEFT);
		HSLFTextRun runDate = dateP.getTextRuns().get(0);
		runDate.setFontSize(12.);
		runDate.setFontFamily("Century Schoolbook");
		runDate.setFontColor(Color.WHITE);
		slide.addShape(date);
	}

	public void slideInformation(HSLFSlideShow ppt, HSLFSlide slide, ArrayList<DBObject> list, int a)
			throws IOException {
		setBackground(slide, version);
		judul(slide, "Information", 1);
		HSLFTextBox content = new HSLFTextBox();
		content.setAnchor(new java.awt.Rectangle(5, 80, 630, 420));
		content.setLineColor(Color.BLACK);
		content.setFillColor(new Color(0, 134, 61));
		content.setLineWidth(2);

		slide.addShape(content);

		// ------------- add picture -----------//
		HSLFPictureData pd3 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\photoTWakun.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3 = new HSLFPictureShape(pd3);
		pictNew3.setAnchor(new java.awt.Rectangle(670, 110, 250, 250));
		pictNew3.setLineColor(new Color(0, 134, 61));
		pictNew3.setLineWidth(3);
		slide.addShape(pictNew3);

		// -------------------------Content------------------------//

		SimpleDateFormat dateFormat = new SimpleDateFormat("dd MMMM yyyy", Locale.ENGLISH);
		dateFormat.setTimeZone(TimeZone.getTimeZone("UTC"));
		String time1 = dateFormat.format(list.get(a).get("user_created_at"));

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Basic Data ");
		content1.appendText("\n\nFull Name:  " + list.get(a).get("user_full_name"), false);
		content1.appendText("\nUsername:  " + "@" + list.get(a).get("username"), false);
		content1.appendText("\nBiography:  " + list.get(a).get("user_description") + ".", false);
		content1.appendText("\nLocation:  - ", false);
		content1.appendText("\nJoin Date:  " + time1, false);
		content1.setAnchor(new java.awt.Rectangle(8, 90, 310, 460));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.JUSTIFY);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(20.);
		runContent1.setFontFamily("Agency FB");
		setTextColor(runContent1, version);
		slide.addShape(content1);

		// -------------------------Content2------------------------//

		HSLFTextBox content2 = new HSLFTextBox();
		content2.setText("Basic Network Info");
		content2.appendText("\n\nFollowings Count: " + list.get(a).get("user_friends_count") + " followings", false);
		content2.appendText("\nFollowers Count: " + list.get(a).get("user_followers_count") + " followers", false);
		content2.appendText("\nTweets Count: " + list.get(a).get("user_statuses_count") + " tweets", false);
		content2.appendText("\nLikes Count:  " + list.get(a).get("user_favourites_count") + " likes", false);
		content2.setAnchor(new java.awt.Rectangle(340, 90, 300, 460));

		HSLFTextParagraph contentP2 = content2.getTextParagraphs().get(0);
		contentP2.setAlignment(TextAlign.JUSTIFY);
		contentP2.setSpaceAfter(2.);
		contentP2.setSpaceBefore(10.);
		contentP2.setLineSpacing(110.0);

		HSLFTextRun runContent2 = contentP2.getTextRuns().get(0);
		runContent2.setFontSize(20.);
		runContent2.setFontFamily("Agency FB");
		runContent2.setFontColor(Color.BLACK);
		slide.addShape(content2);

	}

	public void slideTrendTweet(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String tanggal_puncak,
			long timeframe, int jml_tweet) throws IOException {
		setBackground(slide, version);
		judul(slide, "Trend Tweet", 1);
		String jml_percakapan = "500";

		// ------------- add picture -----------//
		HSLFPictureData pd3a = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\TimeChartTWAccount.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3a = new HSLFPictureShape(pd3a);
		pictNew3a.setAnchor(new java.awt.Rectangle(10, 100, 570, 280));
		slide.addShape(pictNew3a);

		// =====content1=====//
		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Berdasarkan trend graph disamping, dapat dilihat bahwa akun  ");
		content1.appendText(
				" aktif membuat tweet pada  " + tanggal_puncak + ". Data trend tersebut diperoleh dari " + " tweet ",
				false);
		content1.appendText("pada tanggal " + time1 + " hingga " + time2, false);

		content1.setAnchor(new java.awt.Rectangle(600, 100, 350, 390));
		content1.setLineColor(new Color(0, 134, 61));
		content1.setLineWidth(2);
		content1.setVerticalAlignment(VerticalAlignment.MIDDLE);

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.JUSTIFY);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(24.);
		runContent1.setFontFamily("Agency FB");
		setTextColor(runContent1, version);
		slide.addShape(content1);

		// ============content2===========//
		HSLFTextBox content12 = new HSLFTextBox();
		content12.setText("Total Tweet : " + jml_tweet + " tweets");
		content12.appendText("\nTimeframe : " + timeframe + " days", false);

		content12.setAnchor(new java.awt.Rectangle(10, 430, 240, 55));
		content12.setLineColor(new Color(0, 134, 61));
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
		HSLFTextBox content123 = new HSLFTextBox();
		content123.setText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + time2 + ".");

		content123.setAnchor(new java.awt.Rectangle(10, 495, 560, 30));
		content123.setLineColor(new Color(0, 134, 61));
		content123.setLineWidth(2);

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(2.);
		contentP123.setSpaceBefore(10.);
		contentP123.setLineSpacing(110.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(14.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);

	}

	public void slideHashtagCloud(HSLFSlideShow ppt, HSLFSlide slide, List<WordFrequency> wc, String time2)
			throws IOException {
		setBackground(slide, version);
		judul(slide, "Hashtag cloud", 7);

		// ------------- add picture -----------//
		HSLFPictureData pd3e = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\wchtgTWaccount.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3e = new HSLFPictureShape(pd3e);
		pictNew3e.setAnchor(new java.awt.Rectangle(510, 50, 250, 250));
		slide.addShape(pictNew3e);

		// -------------------------Content2------------------------//
		// =====content1=====//
		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Tagar yang paling sering digunakan adalah \"" + wc.get(0).getWord() + "\", \""
				+ wc.get(1).getWord() + "\" dan \"" + wc.get(2).getWord() + "\".");

		content1.setAnchor(new java.awt.Rectangle(340, 330, 530, 130));
		content1.setLineColor(Color.BLACK);
		content1.setLineWidth(3);
		content1.setVerticalAlignment(VerticalAlignment.MIDDLE);

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.CENTER);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(18.);
		runContent1.setFontFamily("Corbel (Body)");
		setTextColor(runContent1, version);
		slide.addShape(content1);

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();
		content123.setText("Ket: Besar tulisan menunjukan tingkat frekuensi kata/hashtag tersebut di tweet.");
		content123.appendText(
				"\nKet: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + time2 + ".", false);
		content123.setAnchor(new java.awt.Rectangle(10, 488, 580, 55));

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(2.);
		contentP123.setSpaceBefore(10.);
		contentP123.setLineSpacing(110.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(14.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);

	}

	public void slideWordcloud(HSLFSlideShow ppt, HSLFSlide slide, List<WordFrequency> wc, String time2)
			throws IOException {
		setBackground(slide, version);
		judul(slide, "Wordcloud", 7);

		// ------------- add picture -----------//
		HSLFPictureData pd31 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\wcTWaccount.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew31 = new HSLFPictureShape(pd31);
		pictNew31.setAnchor(new java.awt.Rectangle(510, 50, 250, 250));
		slide.addShape(pictNew31);

		// -------------------------Content2------------------------//
		// =====content1=====//
		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Kata yang paling sering digunakan adalah \"" + wc.get(0).getWord() + "\", \""
				+ wc.get(1).getWord() + "\" dan \"" + wc.get(2).getWord() + "\".");

		content1.setAnchor(new java.awt.Rectangle(340, 330, 530, 130));
		content1.setLineColor(Color.BLACK);
		content1.setLineWidth(3);
		content1.setVerticalAlignment(VerticalAlignment.MIDDLE);

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.CENTER);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(18.);
		runContent1.setFontFamily("Corbel (Body)");
		setTextColor(runContent1, version);
		slide.addShape(content1);

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();
		content123.setText("Ket: Besar tulisan menunjukan tingkat frekuensi kata/hashtag tersebut di tweet.");
		content123.appendText(
				"\nKet: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + time2 + ".", false);
		content123.setAnchor(new java.awt.Rectangle(10, 488, 580, 55));

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(2.);
		contentP123.setSpaceBefore(10.);
		contentP123.setLineSpacing(110.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(14.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);

	}

	public void slideTweetTypes(HSLFSlide slide, String time2, String jenis_tweet_tertinggi,
			float persentase_tertinggi) {
		judul(slide, "Tweet Types ", 1);

		// -------------------------Content2------------------------ //

		HSLFTextBox content1 = new HSLFTextBox();

		content1.setText("Tipe tweet yang paling banyak di share adalah " + jenis_tweet_tertinggi
				+ " dengan persentase sebesar " + persentase_tertinggi + "%");

		content1.setAnchor(new java.awt.Rectangle(10, 90, 410, 350));
		content1.setLineColor(new Color(0, 134, 61));
		content1.setLineWidth(2);
		content1.setVerticalAlignment(VerticalAlignment.MIDDLE);

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

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();
		content123.setText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + time2 + ".");

		content123.setAnchor(new java.awt.Rectangle(10, 495, 560, 30));
		content123.setLineColor(new Color(0, 134, 61));
		content123.setLineWidth(2);

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(2.);
		contentP123.setSpaceBefore(10.);
		contentP123.setLineSpacing(110.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(14.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);

	}

	public void slideTopPostbyRetweet(HSLFSlide slide, String time2, String nama_akun) {

		judul(slide, "Top Post by Retweet ", 1);

		// -------------------------Content2------------------------ //

		HSLFTextBox content1 = new HSLFTextBox();

		content1.setText("Berikut adalah 5 post yang paling banyak di retweet dari akun " + nama_akun);

		content1.setAnchor(new java.awt.Rectangle(20, 70, 600, 50));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.CENTER);
		contentP1.setSpaceAfter(0.);
		contentP1.setSpaceBefore(0.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(18.);
		runContent1.setFontFamily("Calibri");
		runContent1.setBold(false);
		setTextColor(runContent1, version);

		slide.addShape(content1);

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();
		content123.setText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + time2 + ".");

		content123.setAnchor(new java.awt.Rectangle(10, 495, 560, 30));
		content123.setLineColor(new Color(0, 134, 61));
		content123.setLineWidth(2);

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(2.);
		contentP123.setSpaceBefore(10.);
		contentP123.setLineSpacing(110.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(14.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);

	}

	public void slideTopPostbyFavorite(HSLFSlide slide, String time2, String nama_akun) {

		judul(slide, "Top Post by Favorite ", 1);

		// -------------------------Content2------------------------ //

		HSLFTextBox content1 = new HSLFTextBox();

		content1.setText("Berikut adalah 5 post yang paling banyak di-like dari akun " + nama_akun);

		content1.setAnchor(new java.awt.Rectangle(20, 70, 600, 50));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.CENTER);
		contentP1.setSpaceAfter(0.);
		contentP1.setSpaceBefore(0.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(18.);
		runContent1.setFontFamily("Calibri");
		runContent1.setBold(false);
		setTextColor(runContent1, version);

		slide.addShape(content1);

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();
		content123.setText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + time2 + ".");

		content123.setAnchor(new java.awt.Rectangle(10, 495, 560, 30));
		content123.setLineColor(new Color(0, 134, 61));
		content123.setLineWidth(2);

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(2.);
		contentP123.setSpaceBefore(10.);
		contentP123.setLineSpacing(110.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(14.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);

	}

	public void slideTopMentionAccount(HSLFSlideShow ppt, HSLFSlide slide, ArrayList<String> rankmention, String time2,
			String nama_akun, int banyak_mention) throws IOException {
		judul(slide, "Top Mention Account", 1);

		// ------------- add picture -----------//
		HSLFPictureData pd14 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barmentionTWaccount.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew14 = new HSLFPictureShape(pd14);
		pictNew14.setAnchor(new java.awt.Rectangle(445, 100, 490, 370));
		slide.addShape(pictNew14);

		// -------------------------Content2------------------------ //

		HSLFTextBox content1 = new HSLFTextBox();

		content1.setText("Akun yang paling sering disebut oleh " + nama_akun + " adalah " + rankmention.get(0)
				+ " sebanyak " + banyak_mention + " kali.");

		content1.setAnchor(new java.awt.Rectangle(40, 90, 360, 350));
		content1.setLineColor(new Color(0, 134, 61));
		content1.setLineWidth(2);
		content1.setVerticalAlignment(VerticalAlignment.MIDDLE);

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

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();
		content123.setText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + time2 + ".");

		content123.setAnchor(new java.awt.Rectangle(10, 495, 560, 30));
		content123.setLineColor(new Color(0, 134, 61));
		content123.setLineWidth(2);

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(2.);
		contentP123.setSpaceBefore(10.);
		contentP123.setLineSpacing(110.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(14.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);

	}

	public void slideTopPostbyReply(HSLFSlide slide, String time2, String nama_akun) {

		judul(slide, "Top Post by Reply", 1);

		// -------------------------Content2------------------------ //

		HSLFTextBox content1 = new HSLFTextBox();

		content1.setText("Berikut adalah 5 post yang paling banyak di reply dari akun " + nama_akun);

		content1.setAnchor(new java.awt.Rectangle(20, 70, 600, 50));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.CENTER);
		contentP1.setSpaceAfter(0.);
		contentP1.setSpaceBefore(0.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(18.);
		runContent1.setFontFamily("Calibri");
		runContent1.setBold(false);
		setTextColor(runContent1, version);

		slide.addShape(content1);

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();
		content123.setText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal " + time2 + ".");

		content123.setAnchor(new java.awt.Rectangle(10, 495, 560, 30));
		content123.setLineColor(new Color(0, 134, 61));
		content123.setLineWidth(2);

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(2.);
		contentP123.setSpaceBefore(10.);
		contentP123.setLineSpacing(110.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(14.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);

	}

	public void slideOptimalTimeDiagram(HSLFSlide slide) {
		judul(slide, "Optimal Time Diagram", 1);

		// -------------------------Content2------------------------ //

		HSLFTextBox content1 = new HSLFTextBox();

		content1.setText(
				"Berikut adalah heat map dari akun [account]. Akun [account] mendapat lebih banyak membuat post saat pukul [jam_terbanyak_ngepos] pada hari [hari_terbanyak_ngepos] dengan total [total_post_terbanyak] post.");

		content1.setAnchor(new java.awt.Rectangle(40, 90, 360, 350));
		content1.setLineColor(new Color(0, 134, 61));
		content1.setLineWidth(2);
		content1.setVerticalAlignment(VerticalAlignment.MIDDLE);

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

		// ============content3===========//
		HSLFTextBox content123 = new HSLFTextBox();
		content123.setText("Ket: Berdasarkan data yang disediakan oleh twitter, diakses pada tanggal [dd/mm/yyyy]");

		content123.setAnchor(new java.awt.Rectangle(10, 495, 560, 30));
		content123.setLineColor(new Color(0, 134, 61));
		content123.setLineWidth(2);

		HSLFTextParagraph contentP123 = content123.getTextParagraphs().get(0);
		contentP123.setAlignment(TextAlign.JUSTIFY);
		contentP123.setSpaceAfter(2.);
		contentP123.setSpaceBefore(10.);
		contentP123.setLineSpacing(110.0);

		HSLFTextRun runContent123 = contentP123.getTextRuns().get(0);
		runContent123.setFontSize(14.);
		runContent123.setFontFamily("Corbel (Body)");
		setTextColor(runContent123, version);
		slide.addShape(content123);

	}

	public void slideConclusion(HSLFSlide slide, String time1, String time2) {
		judul(slide, "Kesimpulan", 7);

		// -------------------------Content1------------------------ //

		HSLFTextBox content1 = new HSLFTextBox();

		content1.setText("Berdasarkan data sampel yang disediakan oleh Twitter pada tanggal " + time1 + " hingga "
				+ time2 + ", diakses pada " + time2 + " dapat disimpulkan bahwa:");
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
