package com.ebdesk.report.slide;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.sql.Date;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.text.SimpleDateFormat;
import java.util.Locale;
import java.util.TimeZone;

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
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.VerticalAlignment;
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
import com.mongodb.client.model.Filters;

import javafx.application.Application;
import javafx.embed.swing.SwingFXUtils;
import javafx.scene.Scene;
import javafx.scene.SnapshotParameters;
import javafx.scene.chart.PieChart;
import javafx.scene.image.WritableImage;
import javafx.scene.layout.StackPane;
import javafx.stage.Stage;

public class readjson {

	private int version;
	private String filename;
	private int width;
	private int height;
	private String tema;
	private static String db;
	private static String collect;
	private static String mongoHost = "192.168.114.162";
	private static String mongoPort = "27017";
	private static String mongoUsername = "dev";
	private static String mongoPassword = "Rahas!adev20!8";
	private static String mongoSource = "admin";

	public readjson(String filename, int version, int width, int height, String tema, String db, String collect) {
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

		MongoCursor<Document> cursor = collection.find(Filters.gte("pub_day", "20180709")).iterator();//
		ArrayList<Document> listall = new ArrayList<Document>();
		System.out.println(collection.count(Filters.gte("pub_day", "20180709")));
		System.out.println(collection.count());
		ArrayList<Document> listpro = new ArrayList<Document>();
		ArrayList<Document> listkontra = new ArrayList<Document>();

		try {

			while (cursor.hasNext()) {
				listall.add(cursor.next());
			}
		} catch (Exception e) {
			System.out.println("gagal");
		} finally {
			cursor.close();
		}

		// ---------------time-----------------//
		ArrayList<String> settanggal = new ArrayList<String>(); // tanggal
		Map<String, Integer> map = new HashMap<String, Integer>();

		String time1;
		String time2;
		String[] bulan = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };

		int tanggal1 = 1;
		int tanggal_awal = 999999999;
		int index1 = 0;
		int index2 = 0;
		int tanggal_akhir = 0;

		HashSet<String> setID = new HashSet<String>(); // akun yang mentweet

		Map<String, Integer> mapID = new HashMap<String, Integer>();

		// ====================ALL=====================//
		for (int ii = 0; ii < listall.size(); ii++) {

			ArrayList y7 = (ArrayList) listall.get(ii).get("hashtag"); // hahstag
			try {
				if (y7.size() != 0) {
					for (int ii1 = 0; ii1 < y7.size(); ii1++) {

						if (String.valueOf(y7.get(ii1)).equals("2019gantipresiden")
								|| String.valueOf(y7.get(ii1)).equals("2019kitagantipresiden")
								|| String.valueOf(y7.get(ii1)).equals("2019presidenbaru")
								|| String.valueOf(y7.get(ii1)).equals("gantipresiden2019")
								|| String.valueOf(y7.get(ii1)).equals("2019KitaGantiPresiden")) {
							listkontra.add(listall.get(ii));
						} else if (String.valueOf(y7.get(ii1)).equals("2019tetepjokowi")
								|| String.valueOf(y7.get(ii1)).equals("jokowi2periode")
								|| String.valueOf(y7.get(ii1)).equals("2019t3tepjokowi")
								|| String.valueOf(y7.get(ii1)).equals("t3tepjokowi")
								|| String.valueOf(y7.get(ii1)).equals("2019tetapjokowi")) {
							listpro.add(listall.get(ii));
						}

					}
				}
			} catch (Exception e) {
			}

			// ------------time---------//
			SimpleDateFormat dateFormat = new SimpleDateFormat("dd MMMM yyyy", Locale.ENGLISH);
			dateFormat.setTimeZone(TimeZone.getTimeZone("UTC"));
			time1 = dateFormat.format(listall.get(ii).get("created_at"));

			tanggal1 = Integer.valueOf((String) listall.get(ii).get("pub_day"));
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

			// ===================ID================//

			if (listall.get(ii).get("user_id") != null) {
				setID.add(String.valueOf(listall.get(ii).get("user_id")));
				mapID.put(String.valueOf(listall.get(ii).get("user_id")), ii);

			}
			if (listall.get(ii).get("retweeted_user_id") != null) {
				setID.add(String.valueOf(listall.get(ii).get("retweeted_user_id")));
				mapID.put(String.valueOf(listall.get(ii).get("retweeted_user_id")), ii);
			}
			if (listall.get(ii).get("quoted_user_id") != null) {
				setID.add(String.valueOf(listall.get(ii).get("quoted_user_id")));
				mapID.put(String.valueOf(listall.get(ii).get("quoted_user_id")), ii);
			}

		}

		// ===============TANGGAL=============//
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd MMMM yyyy", Locale.ENGLISH);
		dateFormat.setTimeZone(TimeZone.getTimeZone("UTC"));
		time1 = dateFormat.format(listall.get(index1).get("created_at"));
		time2 = dateFormat.format(listall.get(index2).get("created_at"));

		// ==============tabel=============//
		Map<String, Integer> mapretweet = new HashMap<String, Integer>();

		Object[] px = setID.toArray();
		for (int f = 0; f < setID.size(); f++) {

			if (listall.get(Integer.valueOf(mapID.get(String.valueOf(px[f])))).get("shares_count") != null) {
				mapretweet.put(String.valueOf(mapID.get(String.valueOf(px[f]))),
						(Integer) listall.get(Integer.valueOf(mapID.get(String.valueOf(px[f])))).get("shares_count"));
			}
		}

		ArrayList<String> listindexretweet = new ArrayList<String>(sortbyValue(mapretweet));

		// ===========Information========//
		// gambar1
		try {
			saveImage(String.valueOf(listall.get(Integer.valueOf(listindexretweet.get(listindexretweet.size() - 1)))
					.get("user_profile_image_url")).replace("normal", "200x200"), "liker1TWv2.png");
		} catch (Exception e) {
			// saveImage("https://fajarhac.com/wp-content/uploads/2013/10/IMG_0010.jpg",
			// "liker1TWv2.png");
		}

		// gambar2
		try {
			saveImage(String.valueOf(listall.get(Integer.valueOf(listindexretweet.get(listindexretweet.size() - 2)))
					.get("user_profile_image_url")).replace("normal", "200x200"), "liker2TWv2.png");
		} catch (Exception e) {
			// saveImage("https://fajarhac.com/wp-content/uploads/2013/10/IMG_0010.jpg",
			// "liker2TWv2.png");
		}

		// gambar3
		try {
			saveImage(String.valueOf(listall.get(Integer.valueOf(listindexretweet.get(listindexretweet.size() - 3)))
					.get("user_profile_image_url")).replace("normal", "200x200"), "liker3TWv2.png");
		} catch (Exception e) {
			// saveImage("https://fajarhac.com/wp-content/uploads/2013/10/IMG_0010.jpg",
			// "liker3TWv2.png");
		}

		// =================================PRO====================================//
		HashSet<String> setIDpro = new HashSet<String>(); // akun yang mentweet

		Map<String, Integer> mapIDpro = new HashMap<String, Integer>();

		ArrayList<String> listnamapro = new ArrayList<String>();
		ArrayList<String> listuserpro = new ArrayList<String>();
		ArrayList<String> listlinkpro = new ArrayList<String>();
		ArrayList<String> listhtgpro = new ArrayList<String>();

		List<String> myStringpro = new ArrayList<String>();
		float jml_tweetspro = listpro.size();
		int retweetpro = 0;
		int tweetpro = 0;
		int replypro = 0;
		int quotedpro = 0;
		// ------------------timepro------------------//
		ArrayList<String> settanggalpro = new ArrayList<String>(); // tanggal
		Map<String, Integer> mappro = new HashMap<String, Integer>();
		int tanggal1pro = 0;
		String time1pro;

		for (int ii = 0; ii < listpro.size(); ii++) {

			Pattern rex = Pattern.compile("(?<!\\S)\\p{Alpha}+(?!\\S)"); // words
			try {
				Matcher mx = rex.matcher((String) listpro.get(ii).get("text"));
				while (mx.find()) {
					myStringpro.add(mx.group(0));
				}
			} catch (Exception e) {

			}

			ArrayList x7 = (ArrayList) listpro.get(ii).get("hashtag"); // hashtag
			try {
				if (x7.size() != 0) {
					for (int ii1 = 0; ii1 < x7.size(); ii1++) {
						listhtgpro.add("#" + String.valueOf(x7.get(ii1))); // hashtag

					}
				}
			} catch (Exception e) {
			}

			ArrayList y8 = (ArrayList) listpro.get(ii).get("mention"); // mention
			try {
				if (y8.size() != 0) {
					for (int ii1 = 0; ii1 < y8.size(); ii1++) {
						Document x = (Document) y8.get(ii1);
						listnamapro.add("@" + String.valueOf(x.get("username"))); // akun yang di retweet
					}
				}
			} catch (Exception e) {
			}

			listuserpro.add("@" + (String) listpro.get(ii).get("username")); // akun user

			// ------------time---------//

			dateFormat.setTimeZone(TimeZone.getTimeZone("UTC"));
			time1pro = dateFormat.format(listpro.get(ii).get("created_at"));

			tanggal1pro = Integer.valueOf((String) listpro.get(ii).get("pub_day"));
			settanggalpro.add(time1pro);
			mappro.put(time1pro, ii);
			// =========type data========//

			if (String.valueOf(listpro.get(ii).get("type")).equals("retweet")) {
				retweetpro += 1;
			}

			else if (String.valueOf(listpro.get(ii).get("type")).equals("quote")) {
				quotedpro += 1;
			}

			else if (String.valueOf(listpro.get(ii).get("type")).equals("reply")) {
				replypro += 1;
			}

			else if (String.valueOf(listpro.get(ii).get("type")).equals("tweet")) {
				tweetpro += 1;
			}

			// ===================ID================//

			if (listpro.get(ii).get("user_id") != null) {
				setIDpro.add(String.valueOf(listpro.get(ii).get("user_id")));
				mapIDpro.put(String.valueOf(listpro.get(ii).get("user_id")), ii);

			}
			if (listpro.get(ii).get("retweeted_user_id") != null) {
				setIDpro.add(String.valueOf(listpro.get(ii).get("retweeted_user_id")));
				mapIDpro.put(String.valueOf(listpro.get(ii).get("retweeted_user_id")), ii);
			}
			if (listpro.get(ii).get("quoted_user_id") != null) {
				setIDpro.add(String.valueOf(listpro.get(ii).get("quoted_user_id")));
				mapIDpro.put(String.valueOf(listpro.get(ii).get("quoted_user_id")), ii);
			}

		}
		// ==================Jumlah Akun & Grafik================= //
		int jml_akun = setID.size();
		float persen_retweetpro = retweetpro * 100 / jml_tweetspro;
		float persen_quotedpro = quotedpro * 100 / jml_tweetspro;
		float persen_replypro = replypro * 100 / jml_tweetspro;
		float persen_tweetpro = tweetpro * 100 / jml_tweetspro;
		float persentase_tertinggipro = persen_tweetpro;
		String jenis_tweet_tertinggipro = "Tweet";
		// y
		if (persentase_tertinggipro < persen_retweetpro) {
			persentase_tertinggipro = persen_retweetpro;
			jenis_tweet_tertinggipro = "Retweet";
		}
		if (persentase_tertinggipro < persen_quotedpro) {
			jenis_tweet_tertinggipro = "Quoted";
			persentase_tertinggipro = persen_quotedpro;
		}
		if (persentase_tertinggipro < persen_replypro) {
			jenis_tweet_tertinggipro = "Reply";
			persentase_tertinggipro = persen_replypro;
		}

		PieChartReport chart1pro = new PieChartReport();
		chart1pro.ringchart("Tweet", persen_tweetpro, "Reply", persen_replypro, "Quoted", persen_quotedpro, "Retweet",
				persen_retweetpro, "donatTWv3PRO.png", Color.white);

		// =============Time Series Chart===========//
		String g1pro;
		int nilai_bulanpro = 1;
		String[] time3pro;
		final TimeSeries seriespro = new TimeSeries("Random Data");
		int indexmappro;
		HashSet<String> settime1pro = new HashSet<String>(settanggalpro);
		Object[] rpro = settime1pro.toArray();
		for (int oo = 0; oo < settime1pro.size(); oo++) {
			try {
				indexmappro = mappro.get(String.valueOf(rpro[oo]));

				SimpleDateFormat dateFormat2pro = new SimpleDateFormat("dd MMM yyyy", Locale.ENGLISH);
				dateFormat2pro.setTimeZone(TimeZone.getTimeZone("UTC"));
				g1pro = dateFormat2pro.format(listpro.get(indexmappro).get("created_at"));
				time3pro = g1pro.split("\\s");
				for (int bul = 0; bul < 12; bul++) {

					if (String.valueOf(time3pro[1]).equals(String.valueOf(bulan[bul]))) {

						nilai_bulanpro = (bul + 1);
					}
				}
				seriespro.add(new Day(Integer.valueOf(time3pro[2]), nilai_bulanpro, Integer.valueOf(time3pro[5])),
						new Double(Collections.frequency(settanggalpro, String.valueOf(rpro[oo]))));

			} catch (Exception e) {
			}
		}

		final XYDataset datasetpro = (XYDataset) new TimeSeriesCollection(seriespro);
		JFreeChart timechartpro = ChartFactory.createTimeSeriesChart("Trend Tweet for Keyword: " + tema, "Time",
				"Tweet Count", datasetpro, false, false, false);
		timechartpro.setBackgroundPaint(Color.white);
		XYPlot plotpro = (XYPlot) timechartpro.getPlot();
		plotpro.setBackgroundPaint(new Color(255, 255, 255, 0));
		File timeChartpro = new File("TimeChartTWProv3.png");
		ChartUtilities.saveChartAsJPEG(timeChartpro, timechartpro, 580, 370);

		// ==============tabel=============//
		Map<String, Integer> maplikespro = new HashMap<String, Integer>();
		Map<String, Integer> mapretweetpro = new HashMap<String, Integer>();

		Object[] pxpro = setIDpro.toArray();
		for (int f = 0; f < setIDpro.size(); f++) {
			if (listpro.get(Integer.valueOf(mapIDpro.get(String.valueOf(pxpro[f])))).get("likes_count") != null) {
				maplikespro.put(String.valueOf(mapIDpro.get(String.valueOf(pxpro[f]))), (Integer) listpro
						.get(Integer.valueOf(mapIDpro.get(String.valueOf(pxpro[f])))).get("likes_count"));
			}
			if (listpro.get(Integer.valueOf(mapIDpro.get(String.valueOf(pxpro[f])))).get("shares_count") != null) {
				mapretweetpro.put(String.valueOf(mapIDpro.get(String.valueOf(pxpro[f]))), (Integer) listpro
						.get(Integer.valueOf(mapIDpro.get(String.valueOf(pxpro[f])))).get("shares_count"));
			}
		}

		ArrayList<String> listindexlikespro = new ArrayList<String>(sortbyValue(maplikespro));
		ArrayList<String> listindexretweetpro = new ArrayList<String>(sortbyValue(mapretweetpro));

		HSLFTable table1pro = table_(listpro, listindexlikespro, "likes_count", "Likes Count");
		HSLFTable table2pro = table_(listpro, listindexretweetpro, "shares_count", "Shares Count");

		// ===============BarChart==============//

		BarChart chart7 = new BarChart();
		ArrayList<String> rankhtgpro = chart7.bar(listhtgpro, new HashSet<String>(listhtgpro), "barHashtagTWProv3.png",
				"TOP Used Hashtag", 10, 1, "blue");
		ArrayList<String> ranknamapro = chart7.bar(listnamapro, new HashSet<String>(listnamapro),
				"barMentionTWProv3.png", "TOP Mention User", 10, 1, "blue");
		ArrayList<String> rankuserpro = chart7.bar(listuserpro, new HashSet<String>(listuserpro), "barUserTWProv3.png",
				"TOP Active User", 10, 1, "blue");

		// ==============wordcloud============//

		List<String> myStringhtgpro = new ArrayList<String>(listhtgpro);
		System.out.println("mulai wordcloud");
		Wordcloud word = new Wordcloud();
		List<WordFrequency> wcpro = word.readCNN("wcTWProv3.png", myStringpro, "blue");
		List<WordFrequency> wchtgpro = word.readCNN("wchtgTWProv3.png", myStringhtgpro, "blue");

		// ================================KONTRA====================================//
		HashSet<String> setIDkontra = new HashSet<String>(); // akun yang mentweet

		Map<String, Integer> mapIDkontra = new HashMap<String, Integer>();

		ArrayList<String> listnamakontra = new ArrayList<String>();
		ArrayList<String> listuserkontra = new ArrayList<String>();
		ArrayList<String> listlinkkontra = new ArrayList<String>();
		ArrayList<String> listhtgkontra = new ArrayList<String>();

		List<String> myStringkontra = new ArrayList<String>();
		float jml_tweetskontra = listkontra.size();
		int retweetkontra = 0;
		int tweetkontra = 0;
		int replykontra = 0;
		int quotedkontra = 0;
		// ------------------timepro------------------//
		ArrayList<String> settanggalkontra = new ArrayList<String>(); // tanggal
		Map<String, Integer> mapkontra = new HashMap<String, Integer>();
		int tanggal1kontra = 0;
		String time1kontra;

		for (int ii = 0; ii < listkontra.size(); ii++) {

			Pattern rex = Pattern.compile("(?<!\\S)\\p{Alpha}+(?!\\S)"); // words
			try {
				Matcher mx = rex.matcher((String) listkontra.get(ii).get("text"));
				while (mx.find()) {
					myStringkontra.add(mx.group(0));
				}
			} catch (Exception e) {

			}

			ArrayList x7 = (ArrayList) listkontra.get(ii).get("hashtag"); // hashtag
			try {
				if (x7.size() != 0) {
					for (int ii1 = 0; ii1 < x7.size(); ii1++) {
						listhtgkontra.add("#" + String.valueOf(x7.get(ii1))); // hashtag

					}
				}
			} catch (Exception e) {
			}

			ArrayList y8 = (ArrayList) listkontra.get(ii).get("mention"); // mention
			try {
				if (y8.size() != 0) {
					for (int ii1 = 0; ii1 < y8.size(); ii1++) {
						Document x = (Document) y8.get(ii1);
						listnamakontra.add("@" + String.valueOf(x.get("username"))); // akun yang di retweet
					}
				}
			} catch (Exception e) {
			}

			listuserkontra.add("@" + (String) listkontra.get(ii).get("username")); // akun user

			// ------------time---------//

			dateFormat.setTimeZone(TimeZone.getTimeZone("UTC"));
			time1kontra = dateFormat.format(listkontra.get(ii).get("created_at"));

			tanggal1kontra = Integer.valueOf((String) listkontra.get(ii).get("pub_day"));
			settanggalkontra.add(time1kontra);
			mapkontra.put(time1kontra, ii);
			// =========type data========//

			if (String.valueOf(listkontra.get(ii).get("type")).equals("retweet")) {
				retweetkontra += 1;
			}

			else if (String.valueOf(listkontra.get(ii).get("type")).equals("quote")) {
				quotedkontra += 1;
			}

			else if (String.valueOf(listkontra.get(ii).get("type")).equals("reply")) {
				replykontra += 1;
			}

			else if (String.valueOf(listkontra.get(ii).get("type")).equals("tweet")) {
				tweetkontra += 1;
			}

			// ===================ID================//

			if (listkontra.get(ii).get("user_id") != null) {
				setIDkontra.add(String.valueOf(listkontra.get(ii).get("user_id")));
				mapIDkontra.put(String.valueOf(listkontra.get(ii).get("user_id")), ii);

			}
			if (listkontra.get(ii).get("retweeted_user_id") != null) {
				setIDkontra.add(String.valueOf(listkontra.get(ii).get("retweeted_user_id")));
				mapIDkontra.put(String.valueOf(listkontra.get(ii).get("retweeted_user_id")), ii);
			}
			if (listkontra.get(ii).get("quoted_user_id") != null) {
				setIDkontra.add(String.valueOf(listkontra.get(ii).get("quoted_user_id")));
				mapIDkontra.put(String.valueOf(listkontra.get(ii).get("quoted_user_id")), ii);
			}

		}
		// ==================Jumlah Akun & Grafik================= //
		int jml_akunkontra = setIDkontra.size();
		float persen_retweetkontra = retweetkontra * 100 / jml_tweetskontra;
		float persen_quotedkontra = quotedkontra * 100 / jml_tweetskontra;
		float persen_replykontra = replykontra * 100 / jml_tweetskontra;
		float persen_tweetkontra = tweetkontra * 100 / jml_tweetskontra;
		float persentase_tertinggikontra = persen_tweetkontra;
		String jenis_tweet_tertinggikontra = "Tweet";
		// y
		if (persentase_tertinggikontra < persen_retweetkontra) {
			persentase_tertinggikontra = persen_retweetkontra;
			jenis_tweet_tertinggikontra = "Retweet";
		}
		if (persentase_tertinggikontra < persen_quotedkontra) {
			jenis_tweet_tertinggikontra = "Quoted";
			persentase_tertinggikontra = persen_quotedkontra;
		}
		if (persentase_tertinggikontra < persen_replykontra) {
			jenis_tweet_tertinggikontra = "Reply";
			persentase_tertinggikontra = persen_replykontra;
		}

		PieChartReport chart1kontra = new PieChartReport();
		chart1kontra.ringchart("Tweet", persen_tweetkontra, "Reply", persen_replykontra, "Quoted", persen_quotedkontra,
				"Retweet", persen_retweetkontra, "donatTWv3KONTRA.png", Color.white);

		// =============Time Series Chart===========//
		String g1kontra;
		int nilai_bulankontra = 1;
		String[] time3kontra;
		final TimeSeries serieskontra = new TimeSeries("Random Data");
		int indexmapkontra;
		HashSet<String> settime1kontra = new HashSet<String>(settanggalkontra);
		Object[] rkontra = settime1kontra.toArray();
		for (int oo = 0; oo < settime1kontra.size(); oo++) {
			try {
				indexmapkontra = mapkontra.get(String.valueOf(rkontra[oo]));

				SimpleDateFormat dateFormat2kontra = new SimpleDateFormat("dd MMM yyyy", Locale.ENGLISH);
				dateFormat2kontra.setTimeZone(TimeZone.getTimeZone("UTC"));
				g1kontra = dateFormat2kontra.format(listkontra.get(indexmapkontra).get("created_at"));
				time3kontra = g1kontra.split("\\s");
				for (int bul = 0; bul < 12; bul++) {

					if (String.valueOf(time3kontra[1]).equals(String.valueOf(bulan[bul]))) {

						nilai_bulankontra = (bul + 1);
					}
				}
				serieskontra.add(
						new Day(Integer.valueOf(time3kontra[2]), nilai_bulankontra, Integer.valueOf(time3kontra[5])),
						new Double(Collections.frequency(settanggalkontra, String.valueOf(rkontra[oo]))));

			} catch (Exception e) {
			}
		}

		final XYDataset datasetkontra = (XYDataset) new TimeSeriesCollection(serieskontra);
		JFreeChart timechartkontra = ChartFactory.createTimeSeriesChart("Trend Tweet for Keyword: " + tema, "Time",
				"Tweet Count", datasetkontra, false, false, false);
		timechartkontra.setBackgroundPaint(Color.white);
		XYPlot plotkontra = (XYPlot) timechartkontra.getPlot();
		plotpro.setBackgroundPaint(new Color(255, 255, 255, 0));
		File timeChartkontra = new File("TimeChartTWKontrav3.png");
		ChartUtilities.saveChartAsJPEG(timeChartkontra, timechartkontra, 580, 370);

		// ==============tabel=============//
		Map<String, Integer> maplikeskontra = new HashMap<String, Integer>();
		Map<String, Integer> mapretweetkontra = new HashMap<String, Integer>();

		Object[] pxkontra = setIDkontra.toArray();
		for (int f = 0; f < setIDkontra.size(); f++) {
			if (listkontra.get(Integer.valueOf(mapIDkontra.get(String.valueOf(pxkontra[f]))))
					.get("likes_count") != null) {
				maplikeskontra.put(String.valueOf(mapIDkontra.get(String.valueOf(pxkontra[f]))), (Integer) listkontra
						.get(Integer.valueOf(mapIDkontra.get(String.valueOf(pxkontra[f])))).get("likes_count"));
			}
			if (listkontra.get(Integer.valueOf(mapIDkontra.get(String.valueOf(pxkontra[f]))))
					.get("shares_count") != null) {
				mapretweetkontra.put(String.valueOf(mapIDkontra.get(String.valueOf(pxkontra[f]))), (Integer) listkontra
						.get(Integer.valueOf(mapIDkontra.get(String.valueOf(pxkontra[f])))).get("shares_count"));
			}
		}

		ArrayList<String> listindexlikeskontra = new ArrayList<String>(sortbyValue(maplikeskontra));
		ArrayList<String> listindexretweetkontra = new ArrayList<String>(sortbyValue(mapretweetkontra));

		HSLFTable table1kontra = table_(listkontra, listindexlikeskontra, "likes_count", "Likes Count");
		HSLFTable table2kontra = table_(listkontra, listindexretweetkontra, "shares_count", "Shares Count");

		// ===============BarChart==============//

		ArrayList<String> rankhtgkontra = chart7.bar(listhtgkontra, new HashSet<String>(listhtgkontra),
				"barHashtagTWKontrav3.png", "TOP Used Hashtag", 10, 1, "blue");
		ArrayList<String> ranknamakontra = chart7.bar(listnamakontra, new HashSet<String>(listnamakontra),
				"barMentionTWKontrav3.png", "TOP Mention User", 10, 1, "blue");
		ArrayList<String> rankuserkontra = chart7.bar(listuserkontra, new HashSet<String>(listuserkontra),
				"barUserTWKontrav3.png", "TOP Active User", 10, 1, "blue");

		// ==============wordcloud============//

		List<String> myStringhtgkontra = new ArrayList<String>(listhtgkontra);
		System.out.println("mulai wordcloud");

		List<WordFrequency> wckontra = word.readCNN("wcTWKontrav3.png", myStringkontra, "blue");
		List<WordFrequency> wchtgkontra = word.readCNN("wchtgTWKontrav3.png", myStringhtgkontra, "blue");

		// ======================SLIDE=====================//
		HSLFSlideShow ppt = new HSLFSlideShow();
		ppt.setPageSize(new java.awt.Dimension(width, height));

		// Slide 1
		HSLFSlide slide1 = ppt.createSlide();
		slideMaster(slide1);

		// Slide 2
		HSLFSlide slide2 = ppt.createSlide();
		slideInformation(ppt, slide2, time1, time2, listall, listindexretweet);

		// Slide 3
		HSLFSlide slide3 = ppt.createSlide();
		slideTweetTrend(ppt, slide3, time1, time2);

		// Slide 4
		HSLFSlide slide4 = ppt.createSlide();
		slideNumberofAuthorTrend(ppt, slide4, time1, time2);

		// Slide 5
		HSLFSlide slide5 = ppt.createSlide();
		slideTrendComparisonKontra(ppt, slide5, time1, time2);

		// Slide 6
		HSLFSlide slide6 = ppt.createSlide();
		slideTrendComparisonPro(ppt, slide6, time1, time2);

		// Slide 7
		HSLFSlide slide7 = ppt.createSlide();
		slideTypeofPostKontra(ppt, slide7, time1, time2);

		// Slide 8
		HSLFSlide slide8 = ppt.createSlide();
		slideTypeofPostPro(ppt, slide8, time1, time2);

		// Slide 9
		HSLFSlide slide9 = ppt.createSlide();
		slideTopTweetLikesKontra(slide9, time1, time2);
		slide9.addShape(table1kontra);
		table1kontra.moveTo(20, 100);

		// Slide 10
		HSLFSlide slide10 = ppt.createSlide();
		slideTopTweetLikesPro(slide10, time1, time2);
		slide10.addShape(table1pro);
		table1pro.moveTo(20, 100);

		// Slide 91
		HSLFSlide slide91 = ppt.createSlide();
		slideTopTweetRetweetsKontra(slide91, time1, time2);
		slide91.addShape(table2kontra);
		table2kontra.moveTo(20, 100);

		// Slide 101
		HSLFSlide slide101 = ppt.createSlide();
		slideTopTweetRetweetsPro(slide101, time1, time2);
		slide101.addShape(table2pro);
		table2pro.moveTo(20, 100);

		// Slide 11
		HSLFSlide slide11 = ppt.createSlide();
		slideTopikBahasanKontra(ppt, slide11, time1, time2, wckontra, wchtgkontra);

		// Slide 12
		HSLFSlide slide12 = ppt.createSlide();
		slideTopikBahasanPro(ppt, slide12, time1, time2);

		// Slide 13
		HSLFSlide slide13 = ppt.createSlide();
		slideKelompokKontra(ppt, slide13, time1, time2);

		// Slide 14
		HSLFSlide slide14 = ppt.createSlide();
		slideKelompokPro(ppt, slide14, time1, time2);

		// Slide 14.1
		HSLFSlide slide141 = ppt.createSlide();
		slideOverallNetwork(slide141, time1, time2);

		// Slide 15
		HSLFSlide slide15 = ppt.createSlide();
		slideTop10AccountKontra(slide15, time1, time2);

		// Slide 16
		HSLFSlide slide16 = ppt.createSlide();
		slideTop10AccountPro(slide16, time1, time2);

		// Slide 17
		HSLFSlide slide17 = ppt.createSlide();
		slideTop10HashtagKontra(slide17, time1, time2);

		// Slide 18
		HSLFSlide slide18 = ppt.createSlide();
		slideTop10HashtagPro(slide18, time1, time2);

		// Slide 19
		HSLFSlide slide19 = ppt.createSlide();
		slideNetworkofRetweet(slide19, time1, time2);

		// Slide 20
		HSLFSlide slide20 = ppt.createSlide();
		slideTop10Active(ppt, slide20, time1, time2);

		// Slide 21
		HSLFSlide slide21 = ppt.createSlide();
		slideNetworkofMention(slide21, time1, time2);

		// Slide 22
		HSLFSlide slide22 = ppt.createSlide();
		slideKesimpulan(slide22, time1, time2);

		FileOutputStream out = new FileOutputStream(filename);
		ppt.write(out);
		out.close();
		System.out.println("-------------------> Done <-------------------");
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
		List listall = new LinkedList(mapfollower.entrySet());
		// Defined Custom Comparator here
		Collections.sort(listall, new Comparator() {
			public int compare(Object o1, Object o2) {
				return ((Comparable) ((Map.Entry) (o1)).getValue()).compareTo(((Map.Entry) (o2)).getValue());
			}
		});

		HashMap sortedHashMap = new LinkedHashMap();
		for (Iterator it = listall.iterator(); it.hasNext();) {
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

	private HSLFTable table_(ArrayList<Document> listall, ArrayList<String> listindexlikes, String pp, String hh) {

		String headers[] = { "Date", hh, "Text", "Username" };
		String[][] data = new String[6][4];

		int k = 0;

		for (int v = listindexlikes.size() - 1; v > listindexlikes.size() - 6; v--) { // bawah
			int ko = 0;
			k += 1;
			for (int c = 0; c < 4; c++) {
				try {

					if (ko == 0) {

						String g8 = String.valueOf(
								listall.get(Integer.valueOf(listindexlikes.get(v))).get("created_at").toString());
						String[] time7 = g8.split("\\s");

						data[k - 1][ko] = time7[2] + " " + time7[1] + " " + time7[5];
					} else if (ko == 1) {
						data[k - 1][ko] = String.valueOf(listall.get(Integer.valueOf(listindexlikes.get(v))).get(pp));
					} else if (ko == 2) {

						data[k - 1][ko] = String
								.valueOf(listall.get(Integer.valueOf(listindexlikes.get(v))).get("text"))
								.replaceAll("\n", " ");

					} else if (ko == 3) {
						if (listall.get(Integer.valueOf(listindexlikes.get(v))).get("username") == null) {
							data[k - 1][ko] = String
									.valueOf(listall.get(Integer.valueOf(listindexlikes.get(v))).get("page_name"));
						} else {
							data[k - 1][ko] = String
									.valueOf(listall.get(Integer.valueOf(listindexlikes.get(v))).get("username"));
						}
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
				rt.setFontFamily("Calibri");
				rt.setFontSize(14.);

				if (i == 0) {
					cell.getFill().setForegroundColor(new Color(91, 155, 213));
					rt.setBold(true);
					try {
						cell.setText(headers[j]);

					} catch (Exception e) {
					}
				} else if (i % 2 != 0) {
					cell.setFillColor(new Color(210, 222, 239));
					try {
						cell.setText(data[i - 1][j]);
					} catch (Exception e) {
					}
				} else {
					cell.setFillColor(new Color(234, 239, 247));
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
		border.setLineColor(Color.white);
		table1.setAllBorders(border);

		table1.setColumnWidth(0, 100);
		table1.setColumnWidth(1, 100);
		table1.setColumnWidth(2, 620);
		table1.setColumnWidth(3, 100);

		return table1;

	}

	public void Keterangan(HSLFSlide slide, int mode, String time1, String time2) {

		HSLFTextBox contentp = new HSLFTextBox();
		String tanggalawal = time1;
		String tanggalakhir = time2;

		Date now = new Date(System.currentTimeMillis());
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd MMMM yyyy");

		contentp.setText("Ket: Berdasarkan data sampel yang disediakan oleh twitter pada tanggal " + tanggalawal
				+ " sampai" + tanggalawal + ", diakses pada " + dateFormat.format(now) + ".");

		contentp.setAnchor(new java.awt.Rectangle(30, 490, 480, 35));
		contentp.setLineColor(new Color(255, 192, 0));

		HSLFTextParagraph titlePp = contentp.getTextParagraphs().get(0);
		titlePp.setAlignment(TextAlign.JUSTIFY);
		titlePp.setSpaceAfter(0.);
		titlePp.setSpaceBefore(0.);
		titlePp.setLineSpacing(110.0);

		HSLFTextRun runTitlep = titlePp.getTextRuns().get(0);
		runTitlep.setFontSize(12.);
		runTitlep.setFontFamily("Corbel");
		if (mode == 1) {
			runTitlep.setFontColor(Color.white);
		}

		slide.addShape(contentp);
	}

	public void slideMaster(HSLFSlide slide) {
		setBackground(slide, 1);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("SOCIAL");
		title.appendText("NETWORK", true);
		title.appendText("ANALYSIS", true);
		title.setAnchor(new java.awt.Rectangle(80, 140, 500, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.LEFT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);
		titleP.setLineSpacing(85.0);
		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(77.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setItalic(true);
		runTitle.setFontColor(Color.red);

		HSLFTextParagraph titleP1 = title.getTextParagraphs().get(2);
		HSLFTextRun runTitle1 = titleP1.getTextRuns().get(0);
		runTitle1.setFontSize(77.);
		runTitle1.setFontFamily("Century Schoolbook (Headings)");
		runTitle1.setItalic(true);
		runTitle1.setFontColor(Color.WHITE);
		slide.addShape(title);

		// --------------------Investigation--------------------
		HSLFTextBox issue = new HSLFTextBox();
		issue.setText("ANALISIS " + tema);
		issue.setAnchor(new java.awt.Rectangle(80, 380, 900, 50));
		HSLFTextParagraph issueP = issue.getTextParagraphs().get(0);
		HSLFTextRun runIssue = issueP.getTextRuns().get(0);
		runIssue.setFontSize(20.);
		runIssue.setFontFamily("Corbel (Body)");
		runIssue.setItalic(true);
		runIssue.setFontColor(Color.WHITE);
		slide.addShape(issue);

		// -------------------------Date------------------------//
		HSLFTextBox date = new HSLFTextBox();
		Date now = new Date(System.currentTimeMillis());
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
		date.setText(dateFormat.format(now));
		date.setAnchor(new java.awt.Rectangle(80, 500, 200, 50));
		HSLFTextParagraph dateP = date.getTextParagraphs().get(0);
		dateP.setAlignment(TextAlign.LEFT);
		HSLFTextRun runDate = dateP.getTextRuns().get(0);
		runDate.setFontSize(12.);
		runDate.setItalic(true);
		runDate.setFontFamily("Century Schoolbook");
		runDate.setFontColor(Color.WHITE);
		slide.addShape(date);

		// =====================garis====================//
		HSLFTextBox contentx = new HSLFTextBox();
		contentx.setAnchor(new java.awt.Rectangle(0, 490, 400, 0));
		contentx.setLineColor(Color.WHITE);
		contentx.setRotation(180);
		contentx.setLineWidth(2);
		slide.addShape(contentx);

		HSLFTextBox contentx1 = new HSLFTextBox();
		contentx1.setAnchor(new java.awt.Rectangle(50, 100, 0, 440));
		contentx1.setLineColor(Color.WHITE);
		contentx1.setLineWidth(2);
		slide.addShape(contentx1);

	}

	public void slideInformation(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2,
			ArrayList<Document> listall, ArrayList<String> listindexlikes) throws IOException {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Information");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(50.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);

		// ------------- add picture grafik -----------//
		HSLFPictureData pd3r = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\liker1TWv2.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3r = new HSLFPictureShape(pd3r);
		pictNew3r.setAnchor(new java.awt.Rectangle(28, 355, 75, 75));
		pictNew3r.setShapeType(ShapeType.ELLIPSE);
		slide.addShape(pictNew3r);

		// ------------- add picture grafik -----------//
		HSLFPictureData pd3r2 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\liker2TWv2.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3r2 = new HSLFPictureShape(pd3r2);
		pictNew3r2.setAnchor(new java.awt.Rectangle(250, 165, 75, 75));
		pictNew3r2.setShapeType(ShapeType.ELLIPSE);
		slide.addShape(pictNew3r2);

		// ------------- add picture grafik -----------//
		HSLFPictureData pd3r3 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\liker3TWv2.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3r3 = new HSLFPictureShape(pd3r3);
		pictNew3r3.setAnchor(new java.awt.Rectangle(490, 355, 75, 75));
		pictNew3r3.setShapeType(ShapeType.ELLIPSE);
		slide.addShape(pictNew3r3);

		// -------------------------Content------------------------//

		HSLFTextBox content = new HSLFTextBox();

		content.setText("Laporan ini membahas tentang topik \"" + tema + "\" dalam suatu rentang waktu " + time1
				+ " hingga " + time2);

		content.appendText(" yang beredar pada media sosial twitter dengan total data cuitan yang didapatkan sebanyak "
				+ listall.size() + " data tweet.", false);
		content.setAnchor(new java.awt.Rectangle(100, 460, 760, 60));
		content.setLineColor(new Color(0, 32, 96));
		content.setLineWidth(2);
		content.setVerticalAlignment(VerticalAlignment.MIDDLE);

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.CENTER);

		contentP.setLineSpacing(100.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(17.5);
		runContent.setFontFamily("Calibri");

		slide.addShape(content);

		// -------------------------Content------------------------//

		HSLFTextBox content1 = new HSLFTextBox();

		content1.setText(
				String.valueOf(listall.get(Integer.valueOf(listindexlikes.get(listindexlikes.size() - 1))).get("text"))
						.replaceAll("\n", " "));
		content1.setAnchor(new java.awt.Rectangle(87, 265, 385, 95));
		content1.setLineColor(new Color(127, 143, 175));
		content1.setShapeType(ShapeType.RECT.WEDGE_RECT_CALLOUT);
		content1.setVerticalAlignment(VerticalAlignment.MIDDLE);

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.CENTER);
		contentP1.setLineSpacing(100.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(12.);
		runContent1.setFontFamily("Corbel (Body)");

		slide.addShape(content1);

		// -------------------------Content------------------------//

		HSLFTextBox content2 = new HSLFTextBox();

		content2.setText(
				String.valueOf(listall.get(Integer.valueOf(listindexlikes.get(listindexlikes.size() - 2))).get("text"))
						.replaceAll("\n", " "));
		content2.setAnchor(new java.awt.Rectangle(320, 80, 385, 95));
		content2.setLineColor(new Color(127, 143, 175));
		content2.setShapeType(ShapeType.RECT.WEDGE_RECT_CALLOUT);
		content2.setVerticalAlignment(VerticalAlignment.MIDDLE);

		HSLFTextParagraph contentP2 = content2.getTextParagraphs().get(0);
		contentP2.setAlignment(TextAlign.CENTER);
		contentP2.setLineSpacing(100.0);

		HSLFTextRun runContent2 = contentP2.getTextRuns().get(0);
		runContent2.setFontSize(12.);
		runContent2.setFontFamily("Corbel (Body)");

		slide.addShape(content2);

		// -------------------------Content------------------------//

		HSLFTextBox content3 = new HSLFTextBox();

		content3.setText(
				String.valueOf(listall.get(Integer.valueOf(listindexlikes.get(listindexlikes.size() - 3))).get("text"))
						.replaceAll("\n", " "));
		content3.setAnchor(new java.awt.Rectangle(550, 265, 385, 95));
		content3.setLineColor(new Color(127, 143, 175));
		content3.setShapeType(ShapeType.RECT.WEDGE_RECT_CALLOUT);
		content3.setVerticalAlignment(VerticalAlignment.MIDDLE);

		HSLFTextParagraph contentP3 = content3.getTextParagraphs().get(0);
		contentP3.setAlignment(TextAlign.CENTER);
		contentP3.setLineSpacing(100.0);

		HSLFTextRun runContent3 = contentP3.getTextRuns().get(0);
		runContent3.setFontSize(12.);
		runContent3.setFontFamily("Corbel (Body)");

		slide.addShape(content3);

		// -------------------------Content------------------------//

		HSLFTextBox content12 = new HSLFTextBox();

		content12.setText("@" + String
				.valueOf(listall.get(Integer.valueOf(listindexlikes.get(listindexlikes.size() - 1))).get("username")));
		content12.setAnchor(new java.awt.Rectangle(23, 430, 150, 75));

		HSLFTextParagraph contentP12 = content12.getTextParagraphs().get(0);
		contentP12.setAlignment(TextAlign.JUSTIFY);
		contentP12.setLineSpacing(100.0);

		HSLFTextRun runContent12 = contentP12.getTextRuns().get(0);
		runContent12.setFontSize(12.);
		runContent12.setFontFamily("Corbel");
		runContent12.setFontColor(new Color(255, 192, 0));
		slide.addShape(content12);

		// -------------------------Content------------------------//

		HSLFTextBox content22 = new HSLFTextBox();

		content22.setText("@" + String
				.valueOf(listall.get(Integer.valueOf(listindexlikes.get(listindexlikes.size() - 2))).get("username")));
		content22.setAnchor(new java.awt.Rectangle(245, 240, 150, 75));

		HSLFTextParagraph contentP22 = content22.getTextParagraphs().get(0);
		contentP22.setAlignment(TextAlign.JUSTIFY);
		contentP22.setLineSpacing(100.0);

		HSLFTextRun runContent22 = contentP22.getTextRuns().get(0);
		runContent22.setFontSize(12.);
		runContent22.setFontFamily("Corbel");

		runContent22.setFontColor(new Color(255, 192, 0));
		slide.addShape(content22);

		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("@" + String
				.valueOf(listall.get(Integer.valueOf(listindexlikes.get(listindexlikes.size() - 3))).get("username")));
		content23.setAnchor(new java.awt.Rectangle(485, 430, 150, 75));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.JUSTIFY);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(12.);
		runContent23.setFontFamily("Corbel");

		runContent23.setFontColor(new Color(255, 192, 0));
		slide.addShape(content23);

	}

	public void slideTweetTrend(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2) throws IOException {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Tweet Trend");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(50.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\TweetFrequencyperDay.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(20, 90, 630, 290));
		slide.addShape(pictNew);

		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Rata-rata tweet per hari\t:" + 1233);
		content23.appendText("\nKuartil1\t\t\t:" + 1233, false);
		content23.appendText("\nKuartil2\t\t\t:" + 1233, false);
		content23.appendText("\nKuartil3\t\t\t:" + 1233, false);
		content23.appendText("\nTotal Tweet\t\t:" + 1233, false);
		content23.setAnchor(new java.awt.Rectangle(655, 130, 300, 110));
		content23.setLineColor(new Color(69, 69, 255));
		content23.setLineWidth(2);

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.JUSTIFY);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(16.);
		runContent23.setFontFamily("Corbel");

		slide.addShape(content23);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText("Rata-rata tweet per hari\t:" + 1233);
		content231.appendText("\nKuartil1\t\t\t:" + 1233, false);
		content231.appendText("\nKuartil2\t\t\t:" + 1233, false);
		content231.appendText("\nKuartil3\t\t\t:" + 1233, false);
		content231.appendText("\nTotal Tweet\t\t:" + 1233, false);
		content231.setAnchor(new java.awt.Rectangle(655, 250, 300, 110));
		content231.setLineColor(new Color(255, 14, 14));
		content231.setLineWidth(2);

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(16.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);

	}

	public void slideNumberofAuthorTrend(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2)
			throws IOException {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Number of Author Trend");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(50.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\AuthorFrequencyperDay.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(20, 90, 630, 290));
		slide.addShape(pictNew);

		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Rata-rata author per hari\t:" + 1233);
		content23.appendText("\nKuartil1\t\t\t:" + 1233, false);
		content23.appendText("\nKuartil2\t\t\t:" + 1233, false);
		content23.appendText("\nKuartil3\t\t\t:" + 1233, false);
		content23.appendText("\nTotal\t\t\t:" + 1233, false);
		content23.setAnchor(new java.awt.Rectangle(655, 130, 300, 110));
		content23.setLineColor(new Color(69, 69, 255));
		content23.setLineWidth(2);

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.JUSTIFY);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(16.);
		runContent23.setFontFamily("Corbel");

		slide.addShape(content23);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText("Rata-rata author per hari\t:" + 1233);
		content231.appendText("\nKuartil1\t\t\t:" + 1233, false);
		content231.appendText("\nKuartil2\t\t\t:" + 1233, false);
		content231.appendText("\nKuartil3\t\t\t:" + 1233, false);
		content231.appendText("\nTotal\t\t\t:" + 1233, false);
		content231.setAnchor(new java.awt.Rectangle(655, 250, 300, 110));
		content231.setLineColor(new Color(255, 14, 14));
		content231.setLineWidth(2);

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(16.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);

	}

	public void slideTrendComparisonKontra(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2)
			throws IOException {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Trend Comparison");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\kontra_weeklytrendcomparison.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(50, 140, 400, 200));
		slide.addShape(pictNew);

		// ----- add picture-------------//
		HSLFPictureData pd1 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\kontra_weeklytauthorcomparsion.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
		pictNew1.setAnchor(new java.awt.Rectangle(510, 140, 400, 200));
		slide.addShape(pictNew1);

		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Kelompok Kontra");

		content23.setAnchor(new java.awt.Rectangle(50, 55, 870, 50));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.RIGHT);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(18.);
		runContent23.setFontFamily("Century Schoolbook (Headings)");
		runContent23.setItalic(true);
		slide.addShape(content23);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText("Jumlah tweet yang membuat tweet dengan tagar bernada kontra kepada Jokowi pada rentang "
				+ time1 + " hingga " + time2 + " mengalami pertumbuhan -26,84% dibandingkan minggu sebelumnya");
		content231.setAnchor(new java.awt.Rectangle(30, 400, 440, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(14.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// -------------------------Content------------------------//

		HSLFTextBox content2312 = new HSLFTextBox();

		content2312.setText("Jumlah akun yang menggunakan tagar-tagar bernada kontra kepada Jokowi pada rentang "
				+ time1 + " hingga " + time2 + " mengalami pertumbuhan -15,99% dibandingkan minggu sebelumnya");
		content2312.setAnchor(new java.awt.Rectangle(500, 400, 440, 110));

		HSLFTextParagraph contentP2312 = content2312.getTextParagraphs().get(0);
		contentP2312.setAlignment(TextAlign.JUSTIFY);
		contentP2312.setLineSpacing(100.0);

		HSLFTextRun runContent2312 = contentP2312.getTextRuns().get(0);
		runContent2312.setFontSize(14.);
		runContent2312.setFontFamily("Corbel");

		slide.addShape(content2312);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);

	}

	public void slideTrendComparisonPro(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2)
			throws IOException {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Trend Comparison");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\pro_weeklytrendcomparison.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(50, 140, 400, 200));
		slide.addShape(pictNew);

		// ----- add picture-------------//
		HSLFPictureData pd1 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\pro_weeklytauthorcomparsion.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1 = new HSLFPictureShape(pd);
		pictNew1.setAnchor(new java.awt.Rectangle(510, 140, 400, 200));
		slide.addShape(pictNew1);

		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Kelompok Pro");

		content23.setAnchor(new java.awt.Rectangle(50, 55, 870, 50));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.RIGHT);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(18.);
		runContent23.setFontFamily("Century Schoolbook (Headings)");
		runContent23.setItalic(true);
		slide.addShape(content23);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText("Jumlah tweet yang menyertakan tagar bernada dukungan kepada Jokowi pada rentang " + time1
				+ " hingga " + time2 + " mengalami  pertumbuhan 30,86 dibandingkan minggu sebelumnya.");
		content231.setAnchor(new java.awt.Rectangle(30, 400, 440, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(14.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// -------------------------Content------------------------//

		HSLFTextBox content2312 = new HSLFTextBox();

		content2312.setText("Jumlah akun yang menggunakan tagar bernada dukungan kepada Jokowi pada rentang " + time1
				+ " hingga " + time2 + " mengalami pertumbuhan -1,68%  dibandingkan minggu sebelumnya.");
		content2312.setAnchor(new java.awt.Rectangle(500, 400, 440, 110));

		HSLFTextParagraph contentP2312 = content2312.getTextParagraphs().get(0);
		contentP2312.setAlignment(TextAlign.JUSTIFY);
		contentP2312.setLineSpacing(100.0);

		HSLFTextRun runContent2312 = contentP2312.getTextRuns().get(0);
		runContent2312.setFontSize(14.);
		runContent2312.setFontFamily("Corbel");

		slide.addShape(content2312);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);

	}

	public void slideTypeofPostKontra(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2)
			throws IOException {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Type of Post");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);

		// ----- add picture-------------//
		HSLFPictureData pd1 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\kontra_kontraPostTypesPercentages.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
		pictNew1.setAnchor(new java.awt.Rectangle(50, 40, 580, 430));
		slide.addShape(pictNew1);

		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Kelompok Kontra");

		content23.setAnchor(new java.awt.Rectangle(50, 55, 870, 50));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.RIGHT);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(18.);
		runContent23.setFontFamily("Century Schoolbook (Headings)");
		runContent23.setItalic(true);
		slide.addShape(content23);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText("Tweet didominasi oleh aktivitas retweet. Selebihnya berupa reply dan quote");
		content231.setAnchor(new java.awt.Rectangle(655, 150, 300, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(16.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);

	}

	public void slideTypeofPostPro(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2) throws IOException {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Type of Post");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);
		// ----- add picture-------------//
		HSLFPictureData pd1 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\pro_proPostTypesPercentages.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
		pictNew1.setAnchor(new java.awt.Rectangle(50, 40, 580, 430));
		slide.addShape(pictNew1);

		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Kelompok Pro");

		content23.setAnchor(new java.awt.Rectangle(50, 55, 870, 50));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.RIGHT);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(18.);
		runContent23.setFontFamily("Century Schoolbook (Headings)");
		runContent23.setItalic(true);
		slide.addShape(content23);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText("Tweet didominasi oleh aktivitas retweet. Selebihnya berupa reply dan quote");
		content231.setAnchor(new java.awt.Rectangle(655, 150, 300, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(16.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);
	}

	public void slideTopTweetLikesKontra(HSLFSlide slide, String time1, String time2) throws IOException {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Top Tweet Likes");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);

		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Kelompok Kontra");

		content23.setAnchor(new java.awt.Rectangle(50, 55, 870, 50));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.RIGHT);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(18.);
		runContent23.setFontFamily("Century Schoolbook (Headings)");
		runContent23.setItalic(true);
		slide.addShape(content23);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);

	}

	public void slideTopTweetLikesPro(HSLFSlide slide, String time1, String time2) {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Top Tweet Likes");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);

		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Kelompok Pro");

		content23.setAnchor(new java.awt.Rectangle(50, 55, 870, 50));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.RIGHT);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(18.);
		runContent23.setFontFamily("Century Schoolbook (Headings)");
		runContent23.setItalic(true);
		slide.addShape(content23);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);

	}

	public void slideTopTweetRetweetsKontra(HSLFSlide slide, String time1, String time2) {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Top Tweet Retweets");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);

		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Kelompok Kontra");

		content23.setAnchor(new java.awt.Rectangle(50, 55, 870, 50));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.RIGHT);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(18.);
		runContent23.setFontFamily("Century Schoolbook (Headings)");
		runContent23.setItalic(true);
		slide.addShape(content23);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);

	}

	public void slideTopTweetRetweetsPro(HSLFSlide slide, String time1, String time2) {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Top Tweet Retweets");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);

		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Kelompok Pro");

		content23.setAnchor(new java.awt.Rectangle(50, 55, 870, 50));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.RIGHT);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(18.);
		runContent23.setFontFamily("Century Schoolbook (Headings)");
		runContent23.setItalic(true);
		slide.addShape(content23);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);

	}

	public void slideTopikBahasanKontra(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2,
			List<WordFrequency> wc, List<WordFrequency> wchtg) throws IOException {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Topik Bahasan");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\kontra_wordcloud_wordcloud.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(60, 100, 300, 300));
		slide.addShape(pictNew);

		// ----- add picture-------------//
		HSLFPictureData pd1 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\kontra_hashtagcloud_wordcloud.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
		pictNew1.setAnchor(new java.awt.Rectangle(540, 100, 300, 300));
		slide.addShape(pictNew1);

		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Kelompok Kontra");

		content23.setAnchor(new java.awt.Rectangle(50, 55, 870, 50));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.RIGHT);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(18.);
		runContent23.setFontFamily("Century Schoolbook (Headings)");
		runContent23.setItalic(true);
		slide.addShape(content23);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText(wc.get(0).getWord()
				+ " dan tagar-tagar senada sering disebutkan dalam tweet yang mengandung kata Solo, Jokowi, dan topik-topik pilkada. Beberapa tweet berbunyi dukungan agar Jokowi kembali ke Solo.");
		content231.setAnchor(new java.awt.Rectangle(370, 430, 550, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(14.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);
	}

	public void slideTopikBahasanPro(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2)
			throws IOException {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Topik Bahasan");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\pro_wordcloud_wordcloud.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(60, 100, 300, 300));
		slide.addShape(pictNew);

		// ----- add picture-------------//
		HSLFPictureData pd1 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\pro_hashtagcloud_wordcloud.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
		pictNew1.setAnchor(new java.awt.Rectangle(540, 100, 300, 300));
		slide.addShape(pictNew1);

		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Kelompok Pro");

		content23.setAnchor(new java.awt.Rectangle(50, 55, 870, 50));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.RIGHT);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(18.);
		runContent23.setFontFamily("Century Schoolbook (Headings)");
		runContent23.setItalic(true);
		slide.addShape(content23);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText(
				"#2019TetapJokowi sering disebut bersamaan dengan #Jokowi2Periode. Selain topik-topik pilkada, pembicaraan dengan tagar-tagar tersebut juga melibatkan nama ‘tgb’. Hal ini disebabkan tersebarnya berita  dukungan TGB terhadap Jokowidodo sebagai presiden dalam pemilu medatang.");
		content231.setAnchor(new java.awt.Rectangle(370, 430, 550, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(14.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);
	}

	public void slideOverallNetwork(HSLFSlide slide, String time1, String time2) {

		setBackground(slide, 1);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network of Mention");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, 1);

		slide.addShape(title);

		// ========keterangan=========//
		Keterangan(slide, 1, time1, time2);
	}

	public void slideKelompokKontra(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2) throws IOException {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Kelompok Kontra");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\kontra_TopUsedHashtag.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(40, 80, 420, 300));
		slide.addShape(pictNew);

		// ----- add picture-------------//
		HSLFPictureData pd1 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\kontra_TopMentionedUser.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
		pictNew1.setAnchor(new java.awt.Rectangle(500, 80, 420, 300));
		slide.addShape(pictNew1);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText(
				"Hashtag utama dalam kelompok kontra adalah #2019GantiPresiden dan akun yang paling sering disebut oleh kelompok ini adalah @Muslim_Bersatu1 diikuti oleh @Kaedah_Tawakal_, .");
		content231.setAnchor(new java.awt.Rectangle(50, 410, 860, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(14.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);
	}

	public void slideKelompokPro(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2) throws IOException {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Kelompok Pro");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\pro_TopUsedHashtag.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(40, 80, 420, 300));
		slide.addShape(pictNew);

		// ----- add picture-------------//
		HSLFPictureData pd1 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\pro_TopMentionedUser.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
		pictNew1.setAnchor(new java.awt.Rectangle(500, 80, 420, 300));
		slide.addShape(pictNew1);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText(
				"Hashtag utama dalam kelompok pro adalah #2019TetapJokowi dan akun yang paling sering disebut oleh kelompok ini adalah @jokowi diikuti oleh @RIzmaWidiono.");
		content231.setAnchor(new java.awt.Rectangle(50, 410, 860, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(14.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);
	}

	public void slideTop10AccountKontra(HSLFSlide slide, String time1, String time2) {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Top 10 Accounts");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);
		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Kelompok Kontra");

		content23.setAnchor(new java.awt.Rectangle(50, 55, 870, 50));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.RIGHT);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(18.);
		runContent23.setFontFamily("Century Schoolbook (Headings)");
		runContent23.setItalic(true);
		slide.addShape(content23);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText("Reach = jumlah tweet (count) X jumlah follower."
				+ "Dilihat dari jumlah follower, akun @detikcom dan @Metro_TV adalah akun yang memiliki paling banyak follower dalam kelompok akun-akun yang menggunakan #2019GantiPresiden dan tagar-tagar sejenis. Dilihat dari nilai reach, akun @kompascom dan @Metro_TV merupakan akun yang memiliki nilai tertinggi.");
		content231.setAnchor(new java.awt.Rectangle(50, 410, 860, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(14.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);
	}

	public void slideTop10AccountPro(HSLFSlide slide, String time1, String time2) {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Top 10 Accounts");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);
		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Kelompok Pro");

		content23.setAnchor(new java.awt.Rectangle(50, 55, 870, 50));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.RIGHT);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(18.);
		runContent23.setFontFamily("Century Schoolbook (Headings)");
		runContent23.setItalic(true);
		slide.addShape(content23);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText(
				"Reach = jumlah tweet (count) X jumlah follower.\nDilihat dari jumlah follower dari setiap akun, @fadjroel dan @NKRITheOne adalah akun yang memiliki paling banyak follower dalam percakapan yang mengandung #2019tetapjokowi dan tagar-tagar sejenis. Dilihat dari nilai reach, akun @infokekinian dan @karakterjuang memiliki nilai tertinggi.");
		content231.setAnchor(new java.awt.Rectangle(50, 410, 860, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(14.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);
	}

	public void slideTop10HashtagKontra(HSLFSlide slide, String time1, String time2) {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Top 10 Hashtag Growth");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);
		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Kelompok Kontra");

		content23.setAnchor(new java.awt.Rectangle(50, 55, 870, 50));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.RIGHT);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(18.);
		runContent23.setFontFamily("Century Schoolbook (Headings)");
		runContent23.setItalic(true);
		slide.addShape(content23);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText("Berikut adalah urutan hashtag berdasarkan pertumbuhan tertinggi pada periode " + time1
				+ " hingga " + time2
				+ " dari minggu sebelumnya (25 Juni – 1 Juli 2018).  Hashtag yang penggunaannya  meningkat paling tajam adalah #2019gantisistem.");
		content231.setAnchor(new java.awt.Rectangle(50, 420, 860, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(14.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);
	}

	public void slideTop10HashtagPro(HSLFSlide slide, String time1, String time2) {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Top 10 Hashtag Growth");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);
		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Kelompok Pro");

		content23.setAnchor(new java.awt.Rectangle(50, 55, 870, 50));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.RIGHT);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(18.);
		runContent23.setFontFamily("Century Schoolbook (Headings)");
		runContent23.setItalic(true);
		slide.addShape(content23);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText("Berikut adalah urutan hashtag berdasarkan pertumbuhan tertinggi pada periode " + time1
				+ " hingga " + time2
				+ "dari minggu sebelumnya (25 Juni – 1 Juli 2018).  Hashtag yang penggunaannya  meningkat paling tajam adalah #jokowi2periode.");
		content231.setAnchor(new java.awt.Rectangle(50, 420, 860, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(14.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);
	}

	public void slideNetworkofRetweet(HSLFSlide slide, String time1, String time2) {

		setBackground(slide, 1);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network of Mention");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, 1);

		slide.addShape(title);

		// ========keterangan=========//
		Keterangan(slide, 1, time1, time2);
	}

	public void slideTop10Active(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2) throws IOException {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Top 10 Active Accounts");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\kontra_TopActiveUser.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(5, 100, 460, 360));
		slide.addShape(pictNew);

		// ----- add picture-------------//
		HSLFPictureData pd1 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\pro_TopActiveUser.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
		pictNew1.setAnchor(new java.awt.Rectangle(490, 100, 460, 360));
		slide.addShape(pictNew1);
		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Kelompok Pro");

		content23.setAnchor(new java.awt.Rectangle(50, 70, 300, 50));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.LEFT);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(18.);
		runContent23.setFontFamily("Century Schoolbook (Headings)");
		runContent23.setItalic(true);
		slide.addShape(content23);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText("Kelompok Kontra");

		content231.setAnchor(new java.awt.Rectangle(520, 70, 300, 50));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(18.);
		runContent231.setFontFamily("Century Schoolbook (Headings)");
		runContent231.setItalic(true);
		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2);
	}

	public void slideNetworkofMention(HSLFSlide slide, String time1, String time2) {

		setBackground(slide, 1);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network of Mention");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, 1);

		slide.addShape(title);

		// ========keterangan=========//
		Keterangan(slide, 1, time1, time2);
	}

	public void slideKesimpulan(HSLFSlide slide, String time1, String time2) {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Kesimpulan");
		title.setAnchor(new java.awt.Rectangle(50, 3, 870, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Century Schoolbook (Headings)");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, version);

		slide.addShape(title);

		// -------------------------Content------------------------//
		Date now = new Date(System.currentTimeMillis());
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText("Berdasarkan data sampel yang disediakan oleh twitter pada tanggal " + time1 + " hingga "
				+ time2 + " diakses pada " + dateFormat.format(now) + ", dengan data sebagai berikut:");
		content231.setAnchor(new java.awt.Rectangle(50, 420, 860, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(14.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

	}

}
