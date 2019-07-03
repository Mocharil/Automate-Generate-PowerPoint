package com.ebdesk.report.slide;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
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
import org.apache.poi.hslf.usermodel.HSLFTable;
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

import com.github.davidmoten.rtree.geometry.Line;
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

public class SNATelegramGroup {

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

	public SNATelegramGroup(String filename, int version, int width, int height, String tema) {
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

		MongoDatabase database = mongoClient.getDatabase("criteria_facebook");
		MongoCollection<Document> collection = database.getCollection("330a59fae963a979e2d96a9ec91e8b7c");

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
		int jml_post = list.size();
		int photo = 0;
		int status = 0;
		int video = 0;
		int link = 0;

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
		ArrayList<String> listuser = new ArrayList<String>();
		ArrayList<String> listlink = new ArrayList<String>();
		List<String> myString = new ArrayList<String>();

		for (int ii = 0; ii < list.size(); ii++) {
			String g8 = (String) list.get(ii).get("text");
			Pattern rex = Pattern.compile("(?<!\\S)\\p{Alpha}+(?!\\S)"); // words
			Pattern re = Pattern.compile("(#\\w+)");
			Pattern rel = Pattern.compile("[a-z]+:\\/\\/[^ \\n]*");

			// link//
			try {
				Matcher mx = rel.matcher(g8);
				while (mx.find()) {

					listlink.add(mx.group(0));
				}
			} catch (Exception e) {

			}

			// wordcloud//
			try {
				Matcher mx = rex.matcher(g8);
				while (mx.find()) {
					myString.add(mx.group(0));
				}
			} catch (Exception e) {

			}
			// hashtag//
			try {
				Matcher m = re.matcher(g8);
				while (m.find()) {
					listhtg.add(m.group(0));
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

			settanggal.add(String.valueOf(time1[2]) + " " + String.valueOf(time1[1]) + " " + String.valueOf(time1[5]));
			map.put(String.valueOf(time1[2]) + " " + String.valueOf(time1[1]) + " " + String.valueOf(time1[5]), ii);

			if (tanggal_awal > tanggal1) {
				tanggal_awal = tanggal1;
				index1 = ii;
			}

			if (tanggal_akhir < tanggal1) {
				tanggal_akhir = tanggal1;
				index2 = ii;
			}

			// ===================piechart==============//

			if (String.valueOf(list.get(ii).get("type")).equals("photo")) {
				photo += 1;
			} else if (String.valueOf(list.get(ii).get("type")).equals("status")) {
				status += 1;
			} else if (String.valueOf(list.get(ii).get("type")).equals("video")) {
				video += 1;
			} else if (String.valueOf(list.get(ii).get("type")).equals("link")) {
				link += 1;
			}

			if (list.get(ii).get("user_id") != null) {
				set.add(String.valueOf(list.get(ii).get("user_id")));
				setID.add(String.valueOf(list.get(ii).get("user_id")));
				mapID.put(String.valueOf(list.get(ii).get("user_id")), ii);
			}

			if (list.get(ii).get("username") != null && list.get(ii).get("page_name") != null) {
				listuser.add("@" + (String) list.get(ii).get("username")); // akun user
			}

			else if (list.get(ii).get("username") != null) {

				listuser.add("@" + (String) list.get(ii).get("username")); // akun user
			} else {
				listuser.add("@" + (String) list.get(ii).get("page_name")); // akun page
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
		float persen_photo = photo * 100 / jml_post;
		float persen_status = status * 100 / jml_post;
		float persen_video = video * 100 / jml_post;
		float persen_link = link * 100 / jml_post;
		float persentase_tertinggi = persen_photo;
		String jenis_post_tertinggi = "Photo";
		// y
		if (persentase_tertinggi < persen_status) {
			persentase_tertinggi = persen_status;
			jenis_post_tertinggi = "Status";
		} else if (persentase_tertinggi < persen_video) {
			persentase_tertinggi = persen_video;
			jenis_post_tertinggi = "Video";
		} else if (persentase_tertinggi < persen_link) {
			persentase_tertinggi = persen_link;
			jenis_post_tertinggi = "Link";
		}

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

		Map<String, Integer> maplikes = new HashMap<String, Integer>();

		ArrayList<String> listindexlikes = new ArrayList<String>();

		Object[] px = setID.toArray();
		for (int f = 0; f < setID.size(); f++) {
			maplikes.put(String.valueOf(mapID.get(String.valueOf(px[f]))),
					(Integer) list.get(Integer.valueOf(mapID.get(String.valueOf(px[f])))).get("likes_count"));
		}
		try {
			Map<String, Integer> mapsort = sortByValues(maplikes); // sorted by value
			Set set2 = mapsort.entrySet();
			Iterator iterator2 = set2.iterator();
			while (iterator2.hasNext()) {
				Map.Entry me2 = (Map.Entry) iterator2.next();
				listindexlikes.add((String) me2.getKey());

			}
		} catch (Exception e) {
			System.out.println("MAAF");
		}

		HashSet<String> sethtg = new HashSet<String>(listhtg);

		// int rata_rata_post = (int) (jml_post / (settime1.size() * 24));
		//
		ArrayList<String> rankhtg = new ArrayList<String>();
		// ArrayList<String> ranknama = new ArrayList<String>();
		ArrayList<String> rankuser = new ArrayList<String>();
		ArrayList<String> ranklink = new ArrayList<String>();
		//
		// HashSet<String> setnama = new HashSet<String>(listnama);
		HashSet<String> setuser = new HashSet<String>(listuser);
		HashSet<String> setlink = new HashSet<String>(listlink);
		//
		BarChart chart7 = new BarChart();
		ranktgl = chart7.bar(settanggal, settime1, null, null, 1, 0, null);
		int jml_post_puncak = Collections.frequency(settanggal, String.valueOf(ranktgl.get(0))) / 24;
		String tanggal_puncak = String.valueOf(ranktgl.get(0));
		//
		rankhtg = chart7.bar(listhtg, sethtg, "barHashtagTeleGroup.png", "TOP Used Hashtag", 10, 1, "blue");

		ranklink = chart7.bar(listlink, setlink, "barLinkTeleGroup.png", "TOP Shared Url", 10, 1, "blue");
		//
		int jml_htg = Collections.frequency(listhtg, rankhtg.get(0));

		int jml_link = Collections.frequency(listlink, ranklink.get(0));

		Wordcloud word = new Wordcloud(); // wordcloud
		List<WordFrequency> wc = word.readCNN("wordcloudTeleGroup.png", myString, "blue");

		// ==================slide===============//
		HSLFSlideShow ppt = new HSLFSlideShow();
		ppt.setPageSize(new java.awt.Dimension(width, height));

		// ----- add picture-------------//
		HSLFPictureData pd41r = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\telegram_logo.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew41r = new HSLFPictureShape(pd41r);
		pictNew41r.setAnchor(new java.awt.Rectangle(25, 503, 30, 30));

		// Slide 1
		HSLFSlide slide1 = ppt.createSlide();
		slideMaster(slide1, time1, time2);
		slide1.addShape(pictNew41r);

		// Slide 2
		HSLFSlide slide2 = ppt.createSlide();
		slidePendahuluan(slide2, time1, time2, jml_post);
		slide2.addShape(pictNew41r);

		// Slide 3
		HSLFSlide slide3 = ppt.createSlide();
		slideDaftarGroup(slide3, time1, time2, jml_post);
		slide3.addShape(pictNew41r);

		// Slide 4
		HSLFSlide slide61 = ppt.createSlide();
		slideTop5Group(slide61);
		slide61.addShape(pictNew41r);

		// ------------- add picture -----------//
		HSLFPictureData pd1t3x = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barHashtagTeleGroup.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1t3x = new HSLFPictureShape(pd1t3x);
		pictNew1t3x.setAnchor(new java.awt.Rectangle(30, 80, 450, 350));
		slide61.addShape(pictNew1t3x);

		// Slide 5
		HSLFSlide slide6 = ppt.createSlide();
		slideTop5Akun(slide6);
		slide6.addShape(pictNew41r);

		// ------------- add picture -----------//
		HSLFPictureData pd1t3 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barHashtagTeleGroup.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1t3 = new HSLFPictureShape(pd1t3);
		pictNew1t3.setAnchor(new java.awt.Rectangle(30, 80, 450, 350));
		slide6.addShape(pictNew1t3);
		slide6.addShape(pictNew41r);

		// Slide 6
		HSLFSlide slide4 = ppt.createSlide();
		slideWordcloud(slide4, wc);
		slide4.addShape(pictNew41r);

		// ----- add picture-------------//
		HSLFPictureData pd41 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\wordcloudTeleGroup.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew41 = new HSLFPictureShape(pd41);
		pictNew41.setAnchor(new java.awt.Rectangle(30, 90, 450, 400));
		slide4.addShape(pictNew41);

		// Slide 7
		HSLFSlide slide5 = ppt.createSlide();
		slideJejaring(slide5);
		slide5.addShape(pictNew41r);

		// Slide 7
		HSLFSlide slide7 = ppt.createSlide();
		slideTop10Hashtag(slide7, rankhtg, jml_htg);
		slide7.addShape(pictNew41r);

		// ------------- add picture -----------//
		HSLFPictureData pd31t3 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barHashtagTeleGroup.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew31t3 = new HSLFPictureShape(pd31t3);
		pictNew31t3.setAnchor(new java.awt.Rectangle(30, 80, 450, 350));
		slide7.addShape(pictNew31t3);

		// Slide 8
		HSLFSlide slide8 = ppt.createSlide();
		slideTop10Link(slide8, ranklink, jml_link);
		slide8.addShape(pictNew41r);

		// ------------- add picture -----------//
		HSLFPictureData pd31t35w = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barLinkTeleGroup.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew31t35w = new HSLFPictureShape(pd31t35w);
		pictNew31t35w.setAnchor(new java.awt.Rectangle(250, 80, 450, 300));
		slide8.addShape(pictNew31t35w);
		slide8.addShape(pictNew41r);

		// Slide 9
		HSLFSlide slide9 = ppt.createSlide();
		slideConclusion(slide9);
		slide9.addShape(pictNew41r);

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

	private void table_(ArrayList<Document> list, ArrayList<String> listindexfollower, String namefile, String pp) {
		int k = 0;
		String[][] data = new String[5][4];
		String headers[] = { "Date", pp, "Text", "Username" };
		for (int v = listindexfollower.size() - 1; v > listindexfollower.size() - 11; v--) { // bawah
			int ko = 0;
			k += 1;
			for (int c = 0; c < 4; c++) {
				try {

					if (ko == 0) {

						String g8 = String.valueOf(
								list.get(Integer.valueOf(listindexfollower.get(v))).get("created_at").toString());
						String[] time7 = g8.split("\\s");

						data[k - 1][ko] = time7[2] + " " + time7[1] + " " + time7[5];
					} else if (ko == 1) {
						data[k - 1][ko] = String.valueOf(list.get(Integer.valueOf(listindexfollower.get(v))).get(pp));
					} else if (ko == 2) {
						data[k - 1][ko] = String
								.valueOf(list.get(Integer.valueOf(listindexfollower.get(v))).get("text"));
					} else if (ko == 3) {
						data[k - 1][ko] = String
								.valueOf(list.get(Integer.valueOf(listindexfollower.get(v))).get("username"));
					}
					ko += 1;
				} catch (Exception e) {
				}
			}
		}

	}

	public void Platform(HSLFSlide slide) {

		HSLFTextBox contentp = new HSLFTextBox();

		contentp.setText("Platform: Telegram");

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
		runTitlep.setFontColor(Color.blue);

		slide.addShape(contentp);

	}

	public void Garis(HSLFSlide slide, int a, int b, int c, int d) {
		HSLFTextBox content = new HSLFTextBox();
		content.setAnchor(new java.awt.Rectangle(a, b, c, d));
		content.setLineColor(Color.gray);
		content.setLineWidth(1);
		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		slide.addShape(content);

	}

	public void judul(HSLFSlide slide, String atas, String bawah) {
		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText(atas);
		title.setAnchor(new java.awt.Rectangle(25, 25, 800, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.LEFT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(22.);
		runTitle.setFontFamily("Calibri");
		runTitle.setFontColor(Color.blue);

		slide.addShape(title);

		// --------------------Investigation--------------------
		HSLFTextBox issue = new HSLFTextBox();
		issue.setText(bawah);
		issue.setAnchor(new java.awt.Rectangle(25, 50, 850, 50));
		HSLFTextParagraph issueP = issue.getTextParagraphs().get(0);
		HSLFTextRun runIssue = issueP.getTextRuns().get(0);
		runIssue.setFontSize(17.);
		runIssue.setFontFamily("Calibri");
		runIssue.setBold(false);
		setTextColor(runIssue, version);
		slide.addShape(issue);

		// =========garis===========//
		Garis(slide, 30, 78, 100, 1);

		// =======content Platform========//
		Platform(slide);

	}

	public void slideMaster(HSLFSlide slide, String[] time1, String[] time2) {
		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("SOCIAL NETWORK ANALYSIS");
		title.setAnchor(new java.awt.Rectangle(10, 240, 950, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.CENTER);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(22.);
		runTitle.setFontFamily("Calibri");
		runTitle.setBold(false);
		runTitle.setUnderlined(false);
		runTitle.setFontColor(Color.blue);

		slide.addShape(title);

		// --------------------Investigation--------------------
		HSLFTextBox issue = new HSLFTextBox();
		issue.setText("Analis Telegram Group " + tema + " : " + time1[2] + " " + time1[1] + " " + time1[5] + " - "
				+ time2[2] + " " + time2[1] + " " + time2[5]);
		issue.setAnchor(new java.awt.Rectangle(10, 270, 950, 50));
		HSLFTextParagraph issueP = issue.getTextParagraphs().get(0);
		issueP.setAlignment(TextAlign.CENTER);
		issueP.setSpaceAfter(0.);
		issueP.setSpaceBefore(0.);

		HSLFTextRun runIssue = issueP.getTextRuns().get(0);
		runIssue.setFontSize(18.);
		runIssue.setFontFamily("Calibri");
		runIssue.setBold(false);
		runIssue.setFontColor(Color.GRAY);
		slide.addShape(issue);

		// =========garis===========//
		Garis(slide, 430, 308, 100, 1);

		// =======content Platform========//
		Platform(slide);

	}

	public void slidePendahuluan(HSLFSlide slide, String[] time1, String[] time2, int jml_tweet) {
		judul(slide, "Pendahuluan", "Jumlah data, rentang waktu, dan kata kunci");

		// -------------------------Content------------------------
		String key1 = String.format("%s", "HTILanjutkanPerjuangan");
		String key2 = String.format("%s", "7MeiHTIMenang");
		String key3 = String.format("%s", "IslamSelamatkanNegeri");
		String key4 = String.format("%s", "UmatBersamaHTI");
		String key5 = String.format("%s", "HTILayakMenang");
		String nama_channel = String.format("%s", "HTILayakMenang");

		HSLFTextBox content = new HSLFTextBox();

		content.setText("1. Dalam analisis ini, topik " + nama_channel + " tentang " + tema
				+ " terdiri dari beberapa channel, yaitu :");
		content.appendText("\n\t• " + key1, false);
		content.appendText("\n\t• " + key2, false);
		content.appendText("\n\t• " + key3, false);
		content.appendText("\n\t• " + key4, false);
		content.appendText("\n\t• " + key5, false);
		content.appendText("\n\t• " + key1, false);
		content.appendText("\n\t• " + key1, false);
		content.appendText("\n\t• " + key2, false);
		content.appendText("\n\t• " + key3, false);
		content.appendText("\n\t• " + key4, false);
		content.appendText("\n\t• " + key5, false);
		content.appendText("\n\n2. Data diambil pada tanggal " + time1[2] + " " + time1[1] + " " + time1[5]
				+ " sampai tanggal " + time2[2] + " " + time2[1] + " " + time2[5], false);
		content.appendText(" dengan total group yang didapatkan sebanyak " + jml_tweet + " group.", false);
		content.setAnchor(new java.awt.Rectangle(30, 100, 900, 500));

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.LEFT);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(18.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, version);

		slide.addShape(content);

	}

	public void slideDaftarGroup(HSLFSlide slide, String[] time1, String[] time2, int jml_tweet) {
		judul(slide, "Daftar Group Telegram", "Daftar Group Telegram yang akan dianalisis");

	}

	public void slideWordcloud(HSLFSlide slide, List<WordFrequency> wc) throws Exception {
		judul(slide, "Wordcloud Channel " + tema, "Kata - kata yang paling banyak muncul dan penjelasannya");

		// -------------------------Content2------------------------ //

		HSLFTextBox content1 = new HSLFTextBox();

		content1.setText("Keywords yang paling banyak digunakan adalah '" + wc.get(0).getWord()
				+ "' dengan frekuensi penggunaan sebanyak " + wc.get(0).getFrequency() + ".");
		content1.appendText(" Disusul oleh keyword '" + wc.get(1).getWord() + "' dan '" + wc.get(2).getWord() + "'.",
				false);
		content1.setAnchor(new java.awt.Rectangle(500, 90, 400, 400));

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

	public void slideJejaring(HSLFSlide slide) {
		judul(slide, "Jejaring Group Telegram " + tema, "Keterkaitan antar group berdasarkan aktifitas postnya");

	}

	public void slideTop5Akun(HSLFSlide slide) {
		String nama_akun = "@Ariantos";
		judul(slide, "Aktifitas dari Akun " + nama_akun, "Grafik aktifitas dari group");

	}

	public void slideTop5Group(HSLFSlide slide) {
		judul(slide, "5 Group yang Paling Aktif", "Grafik aktifitas dari group");

	}

	public void slideTop10Hashtag(HSLFSlide slide, ArrayList<String> rankhtg, int jml_htg) {
		judul(slide, "Top 10 Hashtag", "Hashtag yang paling banyak digunakan");

		// -------------------------Content------------------------ //
		HSLFTextBox content = new HSLFTextBox();

		content.setText("Dalam topik ini, hashtag " + rankhtg.get(0)
				+ " menjadi hashtag yang paling banyak digunakan oleh warganet dengan total " + jml_htg + " kali.");
		content.appendText("Disusul oleh hashtag " + rankhtg.get(1) + " dan " + rankhtg.get(2) + ".", false);

		content.setAnchor(new java.awt.Rectangle(500, 250, 430, 155));

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(100.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(18.);
		runContent.setFontFamily("Corbel");
		runContent.setBold(false);
		setTextColor(runContent, version);

		slide.addShape(content);

	}

	public void slideTop10Link(HSLFSlide slide, ArrayList<String> ranklink, int jml_link) {
		judul(slide, "Top 10 Url", "Link yang paling banyak dibagikan");

		// -------------------------Content------------------------ //

		HSLFTextBox content = new HSLFTextBox();

		content.appendText(
				"\n\n• Link yang paling banyak dibagikan adalah " + ranklink.get(0) + " yang dibagikan sebanyak "
						+ jml_link + " kali. Disusul oleh link " + ranklink.get(1) + " dan " + ranklink.get(2) + ".",
				false);
		content.setAnchor(new java.awt.Rectangle(60, 350, 880, 155));

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(18.);
		runContent.setFontFamily("Corbel");
		runContent.setBold(false);
		setTextColor(runContent, version);

		slide.addShape(content);

	}

	public void slideConclusion(HSLFSlide slide) {
		judul(slide, "Kesimpulan Analisis", "Analisis Group Telegram " + tema);
		HSLFTable tabel = new HSLFTable(2, 3);
		tabel.setColumnWidth(2, 3);
		tabel.setAnchor(new java.awt.Rectangle(25, 70, 80, 50));
		slide.addShape(tabel);
	}

}
