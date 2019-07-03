package com.ebdesk.report.slide;

import java.awt.Color;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
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
import java.util.stream.Collectors;

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

public class SNAInstagramTopicv2 {

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

	public SNAInstagramTopicv2(String filename, int version, int width, int height, String tema, String db,
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

		int jml_post = list.size();
		int photo = 0;
		int status = 0;
		int video = 0;

		// ---------------time-----------------//
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

		HashSet<String> setID = new HashSet<String>();
		HashSet<String> set = new HashSet<String>();

		Map<String, Integer> mapID = new HashMap<String, Integer>();
		Map<String, Integer> map = new HashMap<String, Integer>();

		ArrayList<String> listuser = new ArrayList<String>();
		ArrayList<String> listlink = new ArrayList<String>();
		ArrayList<String> listhtg = new ArrayList<String>();
		ArrayList<String> settanggal = new ArrayList<String>();

		List<String> myString = new ArrayList<String>();

		for (int ii = 0; ii < list.size(); ii++) {

			Pattern rex = Pattern.compile("(?<!\\S)\\p{Alpha}+(?!\\S)"); // words
			Pattern re = Pattern.compile("(#\\w+)");
			Pattern rel = Pattern.compile("[a-z]+:\\/\\/[^ \\n]*");

			// link//
			try {
				Matcher mx = rel.matcher((String) list.get(ii).get("text"));
				while (mx.find()) {

					listlink.add(mx.group(0));
				}
			} catch (Exception e) {

			}

			// wordcloud//
			try {
				Matcher mx = rex.matcher((String) list.get(ii).get("text"));
				while (mx.find()) {
					myString.add(mx.group(0));
				}
			} catch (Exception e) {

			}
			// hashtag//
			try {
				Matcher m = re.matcher((String) list.get(ii).get("text"));
				while (m.find()) {
					listhtg.add(m.group(0));
				}
			} catch (Exception e) {

			}

			// ==============TIME============//
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

			if (String.valueOf(list.get(ii).get("type")).equals("image")) {
				photo += 1; // image
			} else if (String.valueOf(list.get(ii).get("type")).equals("carousel")) {
				status += 1; // carousel
			} else if (String.valueOf(list.get(ii).get("type")).equals("video")) {
				video += 1;
			}

			// ==================USERID==================//
			if (list.get(ii).get("user_id") != null) {
				set.add(String.valueOf(list.get(ii).get("user_id")));
				setID.add(String.valueOf(list.get(ii).get("user_id")));
				mapID.put(String.valueOf(list.get(ii).get("user_id")), ii);
			}

			// ======================USERNAME================//
			if (list.get(ii).get("username") != null) {
				listuser.add("@" + (String) list.get(ii).get("username")); // akun user
			}

		}

		// ==============TIME=============//
		HashSet<String> settime1 = new HashSet<String>(settanggal);
		String ga = (String) list.get(index1).get("created_at").toString();
		time1 = ga.split("\\s"); // tanggal awal
		String gk = (String) list.get(index2).get("created_at").toString();
		time2 = gk.split("\\s"); // tanggal akhir

		// ==================jumlah akun & Grafik================= //
		int jml_akun = set.size();
		float persen_photo = photo * 100 / jml_post;
		float persen_status = status * 100 / jml_post;
		float persen_video = video * 100 / jml_post;
		float persen_link = 0;
		float persentase_tertinggi = persen_photo;
		String jenis_post_tertinggi = "Image";
		// y
		if (persentase_tertinggi < persen_status) {
			persentase_tertinggi = persen_status;
			jenis_post_tertinggi = "Multi Media";
		} else if (persentase_tertinggi < persen_video) {
			persentase_tertinggi = persen_video;
			jenis_post_tertinggi = "Video";
		} else if (persentase_tertinggi < persen_link) {
			persentase_tertinggi = persen_link;
			jenis_post_tertinggi = "Link";
		}
		
		PieChartReport chart1 = new PieChartReport();
		chart1.ringchart("Image", persen_photo, "Multi Media", persen_status, "Video", persen_video, null, 0, "donatIGv2.png", Color.white);

		// =================TIMESERIES=================//
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
		JFreeChart timechart = ChartFactory.createTimeSeriesChart("Trend Post for Keyword: " + tema, "Time",
				"Post Count", dataset, false, false, false);
		timechart.setBackgroundPaint(Color.white);
		XYPlot plot = (XYPlot) timechart.getPlot();
		plot.setBackgroundPaint(new Color(255, 255, 255, 0));
		File timeChart = new File("TimeChartIGv2.png");
		ChartUtilities.saveChartAsJPEG(timeChart, timechart, 580, 370);

		// ================TABEL============//
		Map<String, Integer> maplikes = new HashMap<>();

		Object[] px = setID.toArray();
		for (int f = 0; f < setID.size(); f++) {
			if (list.get(Integer.valueOf(mapID.get(String.valueOf(px[f])))).get("likes_count") != null) {
				maplikes.put(String.valueOf(mapID.get(String.valueOf(px[f]))),
						(Integer) list.get(Integer.valueOf(mapID.get(String.valueOf(px[f])))).get("likes_count"));
			}
		}

		ArrayList<String> listindexlikes = new ArrayList<String>(sortbyValue(maplikes));
		HSLFTable table1 = table_(list, listindexlikes, "likes_count", "Likes Count");

		// ==================BARCHART=============//
		BarChart chart7 = new BarChart();
		ArrayList<String> ranktgl = chart7.bar(settanggal, settime1, null, null, 1, 0, null);
		ArrayList<String> rankhtg = chart7.bar(listhtg, new HashSet<String>(listhtg), "barHashtagIGv2.png",
				"TOP Used Hashtag", 10, 1, "orange");
		ArrayList<String> rankuser = chart7.bar(listuser, new HashSet<String>(listuser), "barUserIGv2.png",
				"TOP Active User", 10, 1, "orange");
		ArrayList<String> ranklink = chart7.bar(listlink, new HashSet<String>(listlink), "barLinkIGv2.png",
				"TOP Shared Url", 10, 1, "orange");

		int jml_htg = Collections.frequency(listhtg, rankhtg.get(0));
		int jml_user = Collections.frequency(listuser, rankuser.get(0));
		int jml_link = Collections.frequency(listlink, ranklink.get(0));

		// ============WordCloud============//
		Wordcloud word = new Wordcloud(); // wordcloud
		List<WordFrequency> wc = word.readCNN("wordcloudIGv2.png", myString, "orange");

		// ==============DATA===========//
		int jml_post_puncak = Collections.frequency(settanggal, String.valueOf(ranktgl.get(0))) / 24;
		String tanggal_puncak = String.valueOf(ranktgl.get(0));

		// ===========Information========//
		// gambar1
		saveImage(
				String.valueOf(
						list.get(Integer.valueOf(listindexlikes.get(listindexlikes.size() - 1))).get("user_image")),
				"liker1IGv2.png");

		// gambar2
		saveImage(
				String.valueOf(
						list.get(Integer.valueOf(listindexlikes.get(listindexlikes.size() - 2))).get("user_image")),
				"liker2IGv2.png");

		// gambar3
		saveImage(
				String.valueOf(
						list.get(Integer.valueOf(listindexlikes.get(listindexlikes.size() - 3))).get("user_image")),
				"liker3IGv2.png");

		// ==================slide===============//
		HSLFSlideShow ppt = new HSLFSlideShow();
		ppt.setPageSize(new java.awt.Dimension(width, height));

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\ig_logo.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(25, 503, 30, 30));

		// Slide 1
		HSLFSlide slide1 = ppt.createSlide();
		slideMaster(slide1, time1, time2);
		slide1.addShape(pictNew);

		// Slide 2
		HSLFSlide slide21 = ppt.createSlide();
		slideInformation(ppt, slide21, time1, time2, list, listindexlikes);
		slide21.addShape(pictNew);

		// Slide 3
		HSLFSlide slide3 = ppt.createSlide();
		slideFrekuensiCuitan(ppt, slide3, time1, time2, tanggal_puncak, jml_post_puncak, list.size(), jml_akun,
				jenis_post_tertinggi, persentase_tertinggi);
		slide3.addShape(pictNew);

		// Slide 4
		HSLFSlide slide4 = ppt.createSlide();
		slideWordcloud(ppt, slide4, wc);
		slide4.addShape(pictNew);

		// Slide 5
		HSLFSlide slide5 = ppt.createSlide();
		slideJejaring(slide5);
		slide5.addShape(pictNew);

		// Slide 6
		HSLFSlide slide6 = ppt.createSlide();
		slideTop5Likes(slide6);
		slide6.addShape(pictNew);

		slide6.addShape(table1);
		table1.moveTo(20, 100);

		// Slide 7
		HSLFSlide slide7 = ppt.createSlide();
		slideTop10Hashtag(ppt, slide7, rankhtg, jml_htg);
		slide7.addShape(pictNew);

		// Slide 8
		HSLFSlide slide8 = ppt.createSlide();
		slideTop10User(ppt, slide8, rankuser, ranklink, jml_user, jml_link);
		slide8.addShape(pictNew);

		// Slide 9
		HSLFSlide slide9 = ppt.createSlide();
		slideConclusion(slide9, time1, time2);
		slide9.addShape(pictNew);

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

	private HSLFTable table_(ArrayList<Document> list, ArrayList<String> listindexlikes, String pp, String hh) {
		String headers[] = { "Date", hh, "Text", "Username" };
		String[][] data = new String[6][4];

		int k = 0;

		for (int v = listindexlikes.size() - 1; v > listindexlikes.size() - 6; v--) { // bawah
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
						try {

							data[k - 1][ko] = String
									.valueOf(list.get(Integer.valueOf(listindexlikes.get(v))).get("text"))
									.replaceAll("\n", " ").substring(0, 200) + "...";
						} catch (Exception e) {
							data[k - 1][ko] = String
									.valueOf(list.get(Integer.valueOf(listindexlikes.get(v))).get("text"))
									.replaceAll("\n", " ");

						}

					} else if (ko == 3) {

						data[k - 1][ko] = String
								.valueOf(list.get(Integer.valueOf(listindexlikes.get(v))).get("username"));

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
					cell.getFill().setForegroundColor(new Color(209, 64, 13));
					rt.setBold(true);
					try {
						cell.setText(headers[j]);

					} catch (Exception e) {
					}
				} else if (i % 2 != 0) {
					cell.setFillColor(new Color(245, 124, 81));
					try {
						cell.setText(data[i - 1][j]);
					} catch (Exception e) {
					}
				} else {
					cell.setFillColor(new Color(250, 198, 180));
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
		border.setLineWidth(1.0);
		table1.setAllBorders(border);

		// set width of the 1st column
		table1.setColumnWidth(0, 100);
		// set width of the 2nd column
		table1.setColumnWidth(1, 100);
		table1.setColumnWidth(2, 620);
		table1.setColumnWidth(3, 100);

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

	public void Platform(HSLFSlide slide) {

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
		runTitlep.setBold(false);
		runTitlep.setItalic(true);
		runTitlep.setFontColor(new Color(243, 108, 61));

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
		runTitle.setFontColor(new Color(243, 108, 61));

		slide.addShape(title);

		// --------------------Investigation--------------------
		HSLFTextBox issue = new HSLFTextBox();
		issue.setText(bawah);
		issue.setAnchor(new java.awt.Rectangle(25, 50, 930, 50));
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
		runTitle.setFontColor(new Color(243, 108, 61));

		slide.addShape(title);

		// --------------------Investigation--------------------
		HSLFTextBox issue = new HSLFTextBox();
		issue.setText("Analis Topic " + tema + " : " + time1[2] + " " + time1[1] + " " + time1[5] + " - " + time2[2]
				+ " " + time2[1] + " " + time2[5]);
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

	public void slideInformation(HSLFSlideShow ppt, HSLFSlide slide, String[] time1, String[] time2,
			ArrayList<Document> list, ArrayList<String> listindexlikes) throws IOException {
		judul(slide, "Information", "Jumlah data, rentang waktu, dan tweet terkait");

		// ------------- add picture grafik -----------//
		HSLFPictureData pd3r = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\liker1IGv2.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3r = new HSLFPictureShape(pd3r);
		pictNew3r.setAnchor(new java.awt.Rectangle(28, 320, 75, 75));
		pictNew3r.setShapeType(ShapeType.ELLIPSE);
		slide.addShape(pictNew3r);

		// ------------- add picture grafik -----------//
		HSLFPictureData pd3r2 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\liker2IGv2.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3r2 = new HSLFPictureShape(pd3r2);
		pictNew3r2.setAnchor(new java.awt.Rectangle(330, 320, 75, 75));
		pictNew3r2.setShapeType(ShapeType.ELLIPSE);
		slide.addShape(pictNew3r2);

		// ------------- add picture grafik -----------//
		HSLFPictureData pd3r3 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\liker3IGv2.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3r3 = new HSLFPictureShape(pd3r3);
		pictNew3r3.setAnchor(new java.awt.Rectangle(632, 320, 75, 75));
		pictNew3r3.setShapeType(ShapeType.ELLIPSE);
		slide.addShape(pictNew3r3);

		// -------------------------Content------------------------//

		HSLFTextBox content = new HSLFTextBox();

		content.setText("Laporan ini membahas tentang topik \"" + tema + "\" dalam suatu rentang waktu " + time1[2]
				+ " " + time1[1] + " " + time1[5] + " hingga " + time2[2] + " " + time2[1] + " " + time2[5]);

		content.appendText(" yang beredar pada media sosial Instagram dengan total data post yang didapatkan sebanyak "
				+ list.size() + " data post.", false);
		content.setAnchor(new java.awt.Rectangle(45, 430, 875, 65));
		content.setVerticalAlignment(VerticalAlignment.MIDDLE);
		content.setLineColor(new Color(250, 186, 138));

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.CENTER);

		contentP.setLineSpacing(100.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(17.5);
		runContent.setFontFamily("Calibri");

		slide.addShape(content);

		// -------------------------Content------------------------//

		HSLFTextBox content1 = new HSLFTextBox();

		content1.setText("\""
				+ String.valueOf(list.get(Integer.valueOf(listindexlikes.get(listindexlikes.size() - 1))).get("text"))
						.replaceAll("\n", " ").substring(0, 180)
				+ "...\"");
		content1.setAnchor(new java.awt.Rectangle(87, 100, 250, 210));
		content1.setShapeType(ShapeType.RECT.WEDGE_RECT_CALLOUT);
		content1.setVerticalAlignment(VerticalAlignment.MIDDLE);
		content1.setFillColor(new Color(250, 186, 138));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.CENTER);
		contentP1.setLineSpacing(100.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(17.5);
		runContent1.setFontFamily("Calibri");

		slide.addShape(content1);

		// -------------------------Content------------------------//

		HSLFTextBox content2 = new HSLFTextBox();

		content2.setText("\""
				+ String.valueOf(list.get(Integer.valueOf(listindexlikes.get(listindexlikes.size() - 2))).get("text"))
						.replaceAll("\n", " ").substring(0, 180)
				+ "...\"");
		content2.setAnchor(new java.awt.Rectangle(389, 100, 250, 210));
		content2.setShapeType(ShapeType.RECT.WEDGE_RECT_CALLOUT);
		content2.setVerticalAlignment(VerticalAlignment.MIDDLE);
		content2.setFillColor(new Color(250, 186, 138));

		HSLFTextParagraph contentP2 = content2.getTextParagraphs().get(0);
		contentP2.setAlignment(TextAlign.CENTER);
		contentP2.setLineSpacing(100.0);

		HSLFTextRun runContent2 = contentP2.getTextRuns().get(0);
		runContent2.setFontSize(17.5);
		runContent2.setFontFamily("Calibri");

		slide.addShape(content2);

		// -------------------------Content------------------------//

		HSLFTextBox content3 = new HSLFTextBox();

		content3.setText("\""
				+ String.valueOf(list.get(Integer.valueOf(listindexlikes.get(listindexlikes.size() - 3))).get("text"))
						.replaceAll("\n", " ").substring(0, 180)
				+ "...\"");
		content3.setAnchor(new java.awt.Rectangle(692, 100, 250, 210));
		content3.setLineColor(new Color(61, 140, 255));
		content3.setShapeType(ShapeType.RECT.WEDGE_RECT_CALLOUT);
		content3.setVerticalAlignment(VerticalAlignment.MIDDLE);
		content3.setFillColor(new Color(250, 186, 138));

		HSLFTextParagraph contentP3 = content3.getTextParagraphs().get(0);
		contentP3.setAlignment(TextAlign.CENTER);
		contentP3.setLineSpacing(100.0);

		HSLFTextRun runContent3 = contentP3.getTextRuns().get(0);
		runContent3.setFontSize(17.5);
		runContent3.setFontFamily("Calibri");

		slide.addShape(content3);

		// -------------------------Content------------------------//

		HSLFTextBox content12 = new HSLFTextBox();

		content12.setText("@" + String
				.valueOf(list.get(Integer.valueOf(listindexlikes.get(listindexlikes.size() - 1))).get("username")));
		content12.setAnchor(new java.awt.Rectangle(23, 400, 150, 75));

		HSLFTextParagraph contentP12 = content12.getTextParagraphs().get(0);
		contentP12.setAlignment(TextAlign.JUSTIFY);
		contentP12.setLineSpacing(100.0);

		HSLFTextRun runContent12 = contentP12.getTextRuns().get(0);
		runContent12.setFontSize(16.);
		runContent12.setFontFamily("Calibri");
		runContent12.setBold(true);
		slide.addShape(content12);

		// -------------------------Content------------------------//

		HSLFTextBox content22 = new HSLFTextBox();

		content22.setText("@" + String
				.valueOf(list.get(Integer.valueOf(listindexlikes.get(listindexlikes.size() - 2))).get("username")));
		content22.setAnchor(new java.awt.Rectangle(322, 400, 150, 75));

		HSLFTextParagraph contentP22 = content22.getTextParagraphs().get(0);
		contentP22.setAlignment(TextAlign.JUSTIFY);
		contentP22.setLineSpacing(100.0);

		HSLFTextRun runContent22 = contentP22.getTextRuns().get(0);
		runContent22.setFontSize(16.);
		runContent22.setFontFamily("Calibri");
		runContent22.setBold(true);
		slide.addShape(content22);

		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("@" + String
				.valueOf(list.get(Integer.valueOf(listindexlikes.get(listindexlikes.size() - 3))).get("username")));
		content23.setAnchor(new java.awt.Rectangle(624, 400, 150, 75));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.JUSTIFY);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(16.);
		runContent23.setFontFamily("Calibri");
		runContent23.setBold(true);
		slide.addShape(content23);

	}

	public void slideFrekuensiCuitan(HSLFSlideShow ppt, HSLFSlide slide, String[] time1, String[] time2,
			String tanggal_puncak, int jml_tweet_puncak, int jml_data, int jml_akun, String jenis_tweet_tertinggi,
			float persentase_tertinggi) throws Exception {
		judul(slide, "Frekuensi Kiriman Topik " + tema, "Jumlah Postingan per-hari menurut data Instagram");

		// ------------- add picture grafik -----------//
		HSLFPictureData pd3 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\TimeChartIGv2.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3 = new HSLFPictureShape(pd3);
		pictNew3.setAnchor(new java.awt.Rectangle(40, 85, 480, 260));
		slide.addShape(pictNew3);

		// ----- add picture-------------//
		HSLFPictureData pd4 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\donatIGv2.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew4 = new HSLFPictureShape(pd4);
		pictNew4.setAnchor(new java.awt.Rectangle(540, 50, 400, 310));
		slide.addShape(pictNew4);

		// -------------------------Content------------------------ //

		HSLFTextBox content = new HSLFTextBox();

		content.appendText("•Puncak pembicaraan tertinggi dari issue " + tema + " terjadi pada tanggal "
				+ tanggal_puncak + " dengan jumlah pembicaraan mencapai " + jml_tweet_puncak + " post/hours.", false);
		content.appendText(
				"\n•Jumlah pembicaraan issue ini dalam kurun " + time1[2] + " " + time1[1] + " " + time1[5]
						+ " sampai tanggal " + time2[2] + " " + time2[1] + " " + time2[5] + " sebanyak " + jml_data
						+ " post, dengan jumlah massa Instagram yang membicarakan sebanyak " + jml_akun + " akun.",
				false);
		content.appendText("\n•Hasil diagram tersebut menunjukkan bahwa mayoritas penyebaran issue dilakukan melalui "
				+ jenis_tweet_tertinggi + " dengan persentase sebesar " + persentase_tertinggi + " %.", false);

		content.setAnchor(new java.awt.Rectangle(60, 370, 880, 155));

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(16.);
		runContent.setFontFamily("Corbel");
		runContent.setBold(false);
		setTextColor(runContent, version);

		slide.addShape(content);

	}

	public void slideWordcloud(HSLFSlideShow ppt, HSLFSlide slide, List<WordFrequency> wc) throws Exception {
		judul(slide, "Wordcloud Topik " + tema, "Kata - kata yang paling banyak muncul dan penjelasannya");

		// ----- add picture-------------//
		HSLFPictureData pd41 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\wordcloudIGv2.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew41 = new HSLFPictureShape(pd41);
		pictNew41.setAnchor(new java.awt.Rectangle(30, 90, 450, 400));
		slide.addShape(pictNew41);

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
		judul(slide, "Jejaring Topik " + tema, "Keterkaitan antar akun - akun yang membuat kiriman mengenai " + tema
				+ " berdasarkan aktivitas kirimannya");

	}

	public void slideTop5Likes(HSLFSlide slide) {
		judul(slide, "Top 5 Post", "Kiriman dengan jumlah Like Terbanyak");

	}

	public void slideTop5Share(HSLFSlide slide) {
		judul(slide, "Top 5 Post", "Kiriman dengan jumlah Share Terbanyak");

	}

	public void slideTop10Hashtag(HSLFSlideShow ppt, HSLFSlide slide, ArrayList<String> rankhtg, int jml_htg)
			throws IOException {
		judul(slide, "Top 10 Hashtag", "Hashtag yang paling banyak digunakan dan akun yang paling banyak disebutkan");

		// ------------- add picture -----------//
		HSLFPictureData pd31t3 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barHashtagIGv2.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew31t3 = new HSLFPictureShape(pd31t3);
		pictNew31t3.setAnchor(new java.awt.Rectangle(30, 80, 450, 350));
		slide.addShape(pictNew31t3);

		// -------------------------Content------------------------ //
		HSLFTextBox content = new HSLFTextBox();

		content.setText("Dalam topik ini, hashtag " + rankhtg.get(0)
				+ " menjadi hashtag yang paling banyak digunakan oleh warganet dengan total " + jml_htg + " kali.");
		content.appendText(" Disusul oleh hashtag " + rankhtg.get(1) + " dan " + rankhtg.get(2) + ".", false);

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

	public void slideTop10User(HSLFSlideShow ppt, HSLFSlide slide, ArrayList<String> rankuser,
			ArrayList<String> ranklink, int jml_user, int jml_link) throws IOException {
		judul(slide, "Top 10 Active User dan Url",
				"Akun yang aktif membuat kiriman dan Link yang paling banyak dibagikan");

		// ------------- add picture -----------//
		HSLFPictureData pd31t3q = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barUserIGv2.png"), PictureData.PictureType.PNG);
		HSLFPictureShape pictNew31t3q = new HSLFPictureShape(pd31t3q);
		pictNew31t3q.setAnchor(new java.awt.Rectangle(50, 80, 400, 250));
		slide.addShape(pictNew31t3q);

		// ------------- add picture -----------//
		HSLFPictureData pd31t35w = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barLinkIGv2.png"), PictureData.PictureType.PNG);
		HSLFPictureShape pictNew31t35w = new HSLFPictureShape(pd31t35w);
		pictNew31t35w.setAnchor(new java.awt.Rectangle(450, 80, 400, 250));
		slide.addShape(pictNew31t35w);

		// -------------------------Content------------------------ //

		HSLFTextBox content = new HSLFTextBox();

		content.appendText("• Akun " + rankuser.get(0)
				+ " merupakan akun yang paling aktif dengan jumlah post sebanyak " + jml_user
				+ " terkait topik ini. Disusul oleh " + rankuser.get(1) + " dan " + rankuser.get(2) + ".", false);

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

	public void slideConclusion(HSLFSlide slide, String[] time1, String[] time2) {
		judul(slide, "Kesimpulan Analisis", "Analisis Topik " + tema);

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
		setTextColor(runContent1, version);

		slide.addShape(content1);

	}

}
