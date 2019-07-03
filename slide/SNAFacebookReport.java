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

public class SNAFacebookReport {

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

	public SNAFacebookReport(String filename, int version, int width, int height, String tema, String db,
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

		float jml_post = list.size();
		float photo = 0;
		float status = 0;
		float video = 0;
		float link = 0;

		Map<String, Integer> map = new HashMap<String, Integer>();
		ArrayList<String> listhtg = new ArrayList<String>();
		ArrayList<String> listnama = new ArrayList<String>();
		ArrayList<String> listlink = new ArrayList<String>();
		ArrayList<String> settanggal = new ArrayList<String>();

		String g1;
		String[] time1;
		String[] time2;
		String[] bulan = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
		int nilai_bulan = 1;

		int tanggal1 = 0;

		int tanggal_awal = 999999999;
		int index1 = 0;
		int index2 = 0;
		int tanggal_akhir = 0;

		List<String> myString = new ArrayList<String>();
		List<String> myStringmes = new ArrayList<String>();

		for (int ii = 0; ii < list.size(); ii++) {
			// ====================TEXT==================//
			Pattern rel = Pattern.compile("[a-z]+:\\/\\/[^ \\n]*"); // regex link
			try {
				Matcher mx = rel.matcher((String) list.get(ii).get("description")); // link description
				while (mx.find()) {

					listlink.add(mx.group(0));
				}

				Matcher mx1 = rel.matcher((String) list.get(ii).get("message")); // link message
				while (mx1.find()) {

					listlink.add(mx1.group(0));
				}

			} catch (Exception e) {

			}

			Pattern rex = Pattern.compile("(?<!\\S)\\p{Alpha}+(?!\\S)"); // regex words
			try {
				Matcher mx = rex.matcher((String) list.get(ii).get("description"));// wordcloud decription
				while (mx.find()) {
					myString.add(mx.group(0));

				}
				Matcher mx1 = rex.matcher((String) list.get(ii).get("message")); // // wordcloud message
				while (mx1.find()) {
					myStringmes.add(mx1.group(0));
				}

			} catch (Exception e) {

			}

			Pattern re = Pattern.compile("(#\\w+)"); // regex hashtag
			try {
				Matcher m = re.matcher((String) list.get(ii).get("description")); // hashtag description
				while (m.find()) {
					listhtg.add(m.group(0));
				}

				Matcher m1 = re.matcher((String) list.get(ii).get("message")); // hashtag message
				while (m1.find()) {
					listhtg.add(m1.group(0));
				}
			} catch (Exception e) {

			}

			// ======== nama============//

			listnama.add((String) list.get(ii).get("user_full_name")); // user_full_name

			// ================TIME==============//
			g1 = (String) list.get(ii).get("created_at").toString(); // created_at
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

			// ============type post=============//
			if (String.valueOf(list.get(ii).get("type")).equals("photo")) {
				photo += 1;
			}
			if (String.valueOf(list.get(ii).get("type")).equals("status")) {
				status += 1;
			}
			if (String.valueOf(list.get(ii).get("type")).equals("video")) {
				video += 1;
			}
			if (String.valueOf(list.get(ii).get("type")).equals("link")) {
				link += 1;

			}
		}

		// =============tanggal=============//
		String ga = (String) list.get(index1).get("created_at").toString(); // created_at
		time1 = ga.split("\\s"); // tanggal awal
		String gk = (String) list.get(index2).get("created_at").toString(); // created_at
		time2 = gk.split("\\s"); // tanggal akhir

		// =================TimeSeries=============//
		HashSet<String> settime1 = new HashSet<String>(settanggal);
		String[] time3;
		final TimeSeries series = new TimeSeries("Random Data");
		int indexmap;
		Object[] r = settime1.toArray();

		for (int oo = 0; oo < settime1.size(); oo++) {
			try {
				indexmap = map.get(String.valueOf(r[oo]));
				g1 = (String) list.get(indexmap).get("created_at").toString(); // created_at
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
		JFreeChart timechart = ChartFactory.createTimeSeriesChart("Trend Facebook Posts Per Day: " + tema, "Time",
				"Post Count", dataset, false, false, false);
		timechart.setBackgroundPaint(Color.white);
		XYPlot plot = (XYPlot) timechart.getPlot();
		plot.setBackgroundPaint(Color.white);
		int width1 = 580; /* Width of the image */
		int height1 = 370; /* Height of the image */
		File timeChart = new File("TimeChartFBv1.png");
		ChartUtilities.saveChartAsPNG(timeChart, timechart, width1, height1);

		// =================BarChart=============//
		BarChart chart7 = new BarChart();
		ArrayList<String> ranktgl = chart7.bar(settanggal, settime1, null, null, 1, 0, "red");
		ArrayList<String> rankhtg = chart7.bar(listhtg, new HashSet<String>(listhtg), "hashtagFBv1.png",
				"Most Used Hashtag in Post Based on Related Topic", 10, 1, "red");
		ArrayList<String> ranklink = chart7.bar(listlink, new HashSet<String>(listlink), "linkFBv1.png",
				"Most Used Domain in Post Based on Related Topic", 10, 1, "red");
		ArrayList<String> ranknama = chart7.bar(listnama, new HashSet<String>(listnama), "namaFBv1.png",
				"10 Active Accounts Based on Related Topic", 10, 1, "red");

		// ===============GRAFIK TYPE POST================//

		float persen_status = status * 100 / jml_post;
		float persen_video = video * 100 / jml_post;
		float persen_photo = photo * 100 / jml_post;
		float persen_link = link * 100 / jml_post;
		float persentase_tertinggi = persen_status;
		String jenis_post = "Status";

		if (persentase_tertinggi < persen_photo) {
			persentase_tertinggi = persen_photo;
			jenis_post = "Photo";
		}
		if (persentase_tertinggi < persen_video) {
			jenis_post = "Video";
			persentase_tertinggi = persen_video;
		}
		if (persentase_tertinggi < persen_link) {
			jenis_post = "Link";
			persentase_tertinggi = persen_link;
		}

		PieChartReport chart1 = new PieChartReport();
		chart1.ringchart("Photo", persen_photo, "Video", persen_video, "Link", persen_link, "Status", persen_status, "donatFBv1.png", Color.white);
		
		
		// ===============wordcloud================//
		Wordcloud word = new Wordcloud(); // wordcloud
		List<WordFrequency> wcdes = word.readCNN("wcdesFBv1.png", myString, "biru");
		List<WordFrequency> wcmes = word.readCNN("wcmesFBv1.png", myStringmes, "blue");

		// ==============DATA============//
		int post_puncak = Collections.frequency(settanggal, String.valueOf(ranktgl.get(0)));
		String tanggal_puncak = String.valueOf(ranktgl.get(0));
		int rata_rata_post = (int) list.size() / settime1.size();

		// ---------------------------SLIDE--------------------------//

		HSLFSlideShow ppt = new HSLFSlideShow();
		ppt.setPageSize(new java.awt.Dimension(width, height));

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\FB_logo.png"),
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
		slideTrendPost1(ppt, slide3, time1, time2, list.size(), rata_rata_post, tanggal_puncak, post_puncak, wcmes,
				wcdes);
		slide3.addShape(pictNew);

		// Slide 4
		HSLFSlide slide4 = ppt.createSlide();
		slideTipeTautanPost(ppt, slide4, jenis_post, persentase_tertinggi);
		slide4.addShape(pictNew);

		// Slide 5
		HSLFSlide slide5 = ppt.createSlide();
		slideAccountPageGroup(ppt, slide5, ranknama);
		slide5.addShape(pictNew);

		// // Slide 6
		HSLFSlide slide6 = ppt.createSlide();
		slideHashtag(ppt, slide6, rankhtg);
		slide6.addShape(pictNew);

		// // Slide 7
		HSLFSlide slide7 = ppt.createSlide();
		slideLink(ppt, slide7, ranklink);
		slide7.addShape(pictNew);

		// Slide 8
		HSLFSlide slide8 = ppt.createSlide();
		slideNetworkofPost(slide8);
		slide8.addShape(pictNew);

		// ------------- add picturegephi -----------//
		HSLFPictureData pd3gx1 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\gephi2.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3gx1 = new HSLFPictureShape(pd3gx1);
		pictNew3gx1.setAnchor(new java.awt.Rectangle(30, 50, 450, 430));
		slide8.addShape(pictNew3gx1);

		// // Slide 9
		HSLFSlide slide9 = ppt.createSlide();
		slideNetworkofPost2(slide9);
		slide9.addShape(pictNew);

		// ------------- add picturegephi -----------//
		HSLFPictureData pd3gx = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\gephi.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3gx = new HSLFPictureShape(pd3gx);
		pictNew3gx.setAnchor(new java.awt.Rectangle(30, 60, 460, 430));
		slide9.addShape(pictNew3gx);

		// // Slide 11
		HSLFSlide slide11 = ppt.createSlide();
		slideNetworkofShare(slide11);
		slide11.addShape(pictNew);

		// ------------- add picturegephi -----------//
		HSLFPictureData pd3gx2 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\gephi.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3gx2 = new HSLFPictureShape(pd3gx2);
		pictNew3gx2.setAnchor(new java.awt.Rectangle(30, 60, 460, 430));
		slide11.addShape(pictNew3gx2);

		// // Slide 12
		HSLFSlide slide12 = ppt.createSlide();
		slideNetworkofShare2(slide12);
		slide12.addShape(pictNew);

		// ------------- add picturegephi -----------//
		HSLFPictureData pd3gx21 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\gephi.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3gx21 = new HSLFPictureShape(pd3gx21);
		pictNew3gx21.setAnchor(new java.awt.Rectangle(30, 60, 460, 430));
		slide12.addShape(pictNew3gx21);

		// // Slide 13
		HSLFSlide slide13 = ppt.createSlide();
		slideConclution(slide13, time1, time2);
		slide13.addShape(pictNew);

		FileOutputStream out = new FileOutputStream(filename);
		ppt.write(out);
		out.close();
		System.out.println("-------------------> Done <-------------------");
	}

	// ========Platform====//
	public void Platform(HSLFSlide slide) {

		HSLFTextBox contentp = new HSLFTextBox();

		contentp.setText("Platform: Facebook");

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
		setTextColor(runTitle1, version);
		slide.addShape(title);

		// --------------------Investigation--------------------
		HSLFTextBox issue = new HSLFTextBox();

		issue.appendText("INVESTIGASI ISU " + tema, false);
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
		titleP.setAlignment(TextAlign.LEFT);
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
		content.appendText(time2[2] + "/" + time2[1] + "/" + time2[5] + ", issue mengenai '", false);
		content.appendText(tema + "' sedang ramai diperbincangkan di media social facebook.", false);
		content.setAnchor(new java.awt.Rectangle(280, 165, 400, 130));
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

	public void slideTrendPost1(HSLFSlideShow ppt, HSLFSlide slide, String[] time1, String[] time2, int len_json,
			float rata_rata_post, String tanggal_puncak, int jml_post_puncak, List<WordFrequency> wcmes,
			List<WordFrequency> wcdes) throws IOException {
		setBackground(slide, 2);

		HSLFTextBox title = new HSLFTextBox();

		title.setText("Trend Post");

		title.setAnchor(new java.awt.Rectangle(30, 40, 200, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.JUSTIFY);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(38.);
		runTitle.setFontFamily("Times New Roman");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, 2);

		slide.addShape(title);

		// ------------- add picture grafik -----------//
		HSLFPictureData pd3 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\TimeChartFBv1.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3 = new HSLFPictureShape(pd3);
		pictNew3.setAnchor(new java.awt.Rectangle(420, 50, 520, 240));
		slide.addShape(pictNew3);

		// ------------- add picture Description -----------//
		HSLFPictureData pd1 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\wcdesFBv1.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
		pictNew1.setAnchor(new java.awt.Rectangle(480, 310, 180, 180));
		slide.addShape(pictNew1);

		// ------------- add picture capption -----------//
		HSLFPictureData pd2 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\wcmesFBv1.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew2 = new HSLFPictureShape(pd2);
		pictNew2.setAnchor(new java.awt.Rectangle(700, 310, 180, 180));
		slide.addShape(pictNew2);

		// =======content2========//
		HSLFTextBox content2 = new HSLFTextBox();

		content2.setText("\nBerdasarkan trend dari tanggal " + time1[2] + "/" + time1[1] + "/" + time1[5]);
		content2.appendText(" hingga " + time2[2] + "/" + time2[1] + "/" + time2[5], false);
		content2.appendText(" diperoleh jumlah postingan sebanyak ", false);
		content2.appendText(String.valueOf(len_json), false);
		content2.appendText(" post dengan rata-rata post " + rata_rata_post
				+ "/hari, dengan puncak pembahasan terjadi pada tanggal ", false);
		content2.appendText(
				tanggal_puncak + " yaitu mencapai sekitar " + String.valueOf(jml_post_puncak) + " postingan.", false);
		content2.appendText(
				"\n\nKeyword yang sering digunakan berdasarkan postingan yang di share (description) adalah ", false);
		content2.appendText(
				wcdes.get(0).getWord() + ", " + wcdes.get(1).getWord() + " dan " + wcdes.get(2).getWord() + ".", false);
		content2.appendText(
				"\nSedangkan kata-kata yang sering digunakan berdasarkan caption dari postingan yang di-share (Message) adalah ",
				false);
		content2.appendText(
				wcmes.get(0).getWord() + ", " + wcmes.get(1).getWord() + " dan " + wcmes.get(2).getWord() + ".", false);

		content2.setAnchor(new java.awt.Rectangle(10, 100, 350, 385));

		HSLFTextParagraph titleP2 = content2.getTextParagraphs().get(0);
		titleP2.setAlignment(TextAlign.JUSTIFY);
		titleP2.setSpaceAfter(0.);
		titleP2.setSpaceBefore(0.);
		titleP2.setLineSpacing(115.0);

		HSLFTextRun runTitle2 = titleP2.getTextRuns().get(0);
		runTitle2.setFontSize(15.);
		runTitle2.setFontFamily("Calibri");
		runTitle2.setBold(false);
		setTextColor(runTitle2, 2);

		slide.addShape(content2);

		// =======content 3========//
		HSLFTextBox content3 = new HSLFTextBox();

		content3.setText("Ket: ukuran tulisan menunjukan tingkat frekuensi kata tersebut pada postingan.");

		content3.setAnchor(new java.awt.Rectangle(10, 450, 350, 50));
		content3.setLineColor(Color.BLUE);

		HSLFTextParagraph titleP3 = content3.getTextParagraphs().get(0);
		titleP3.setAlignment(TextAlign.JUSTIFY);
		titleP3.setSpaceAfter(0.);
		titleP3.setSpaceBefore(0.);
		titleP3.setLineSpacing(110.0);

		HSLFTextRun runTitle3 = titleP3.getTextRuns().get(0);
		runTitle3.setFontSize(15.);
		runTitle3.setFontFamily("Calibri");
		runTitle3.setBold(false);
		setTextColor(runTitle3, 2);

		slide.addShape(content3);

		// =======content 4========//
		HSLFTextBox content4 = new HSLFTextBox();

		content4.setText("Description");

		content4.setAnchor(new java.awt.Rectangle(525, 500, 90, 30));
		content4.setLineColor(Color.BLUE);

		HSLFTextParagraph titleP4 = content4.getTextParagraphs().get(0);
		titleP4.setAlignment(TextAlign.JUSTIFY);
		titleP4.setSpaceAfter(0.);
		titleP4.setSpaceBefore(0.);
		titleP4.setLineSpacing(110.0);

		HSLFTextRun runTitle4 = titleP4.getTextRuns().get(0);
		runTitle4.setFontSize(15.);
		runTitle4.setFontFamily("Calibri");
		runTitle4.setBold(false);
		setTextColor(runTitle4, 2);

		slide.addShape(content4);

		// =======content 5========//
		HSLFTextBox content5 = new HSLFTextBox();

		content5.setText("Messages");

		content5.setAnchor(new java.awt.Rectangle(750, 500, 90, 30));
		content5.setLineColor(Color.BLUE);

		HSLFTextParagraph titleP5 = content5.getTextParagraphs().get(0);
		titleP5.setAlignment(TextAlign.JUSTIFY);
		titleP5.setSpaceAfter(0.);
		titleP5.setSpaceBefore(0.);
		titleP5.setLineSpacing(110.0);

		HSLFTextRun runTitle5 = titleP5.getTextRuns().get(0);
		runTitle5.setFontSize(15.);
		runTitle5.setFontFamily("Calibri");
		runTitle5.setBold(false);
		setTextColor(runTitle5, 2);

		slide.addShape(content5);

		// =======content Platform========//
		Platform(slide);

	}

	public void slideTipeTautanPost(HSLFSlideShow ppt, HSLFSlide slide, String jenis_post, float persentase_tertinggi)
			throws IOException {

		setBackground(slide, 2);

		HSLFTextBox title = new HSLFTextBox();

		title.setText("Tipe Tautan\nPost");

		title.setAnchor(new java.awt.Rectangle(25, 30, 300, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(45.);
		runTitle.setFontFamily("Times New Roman");
		runTitle.setBold(false);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		setTextColor(runTitle, 2);

		slide.addShape(title);

		// --------------------grafik-------------------//
		HSLFPictureData pd4 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\donatFBv1.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew4 = new HSLFPictureShape(pd4);
		pictNew4.setAnchor(new java.awt.Rectangle(470, 45, 415, 400));
		slide.addShape(pictNew4);

		// =======content2========//
		HSLFTextBox content2 = new HSLFTextBox();

		content2.setText("\nBerdasarkan diagram disamping terlihat bahwa sebagian besar postingan berupa ");
		content2.appendText(jenis_post + " dengan persentase sebesar " + persentase_tertinggi + " %.", false);

		content2.setAnchor(new java.awt.Rectangle(35, 250, 370, 385));

		HSLFTextParagraph titleP2 = content2.getTextParagraphs().get(0);
		titleP2.setAlignment(TextAlign.JUSTIFY);
		titleP2.setSpaceAfter(0.);
		titleP2.setSpaceBefore(0.);
		titleP2.setLineSpacing(110.0);

		HSLFTextRun runTitle2 = titleP2.getTextRuns().get(0);
		runTitle2.setFontSize(20.);
		runTitle2.setFontFamily("Calibri");
		runTitle2.setBold(false);
		setTextColor(runTitle2, 2);

		slide.addShape(content2);

		// =======content Platform========//
		Platform(slide);
	}

	// ------------------------------------------------new---------------------------------------------

	public void slideAccountPageGroup(HSLFSlideShow ppt, HSLFSlide slide, ArrayList<String> rank) throws IOException {
		setBackground(slide, 2);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Account/Page/Group yang Paling\nSering Memposting");
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
		HSLFPictureData pd41 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\namaFBv1.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew41 = new HSLFPictureShape(pd41);
		pictNew41.setAnchor(new java.awt.Rectangle(120, 170, 700, 230));
		slide.addShape(pictNew41);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();

		content.appendText(
				"Berdasarkan diagram diatas, terlihat account/page/group yang paling sering memposting terkait ",
				false);
		content.appendText(tema + " adalah " + rank.get(0) + " diikuti oleh " + rank.get(1) + " dan " + rank.get(2),
				false);

		content.setAnchor(new java.awt.Rectangle(220, 420, 520, 70));

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

	public void slideHashtag(HSLFSlideShow ppt, HSLFSlide slide, ArrayList<String> rank) throws IOException {
		setBackground(slide, 2);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Hashtag yang Paling Sering\nDigunakan");
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
		HSLFPictureData pd412 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\hashtagFBv1.png"), PictureData.PictureType.PNG);
		HSLFPictureShape pictNew412 = new HSLFPictureShape(pd412);
		pictNew412.setAnchor(new java.awt.Rectangle(120, 170, 700, 230));
		slide.addShape(pictNew412);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();

		content.appendText(
				"Berdasarkan diagram diatas, terlihat bahwasanya hashtag yang sering digunakan pada postingan adalah  ",
				false);
		content.appendText(rank.get(0) + " diikuti oleh " + rank.get(1) + " dan " + rank.get(2), false);

		content.setAnchor(new java.awt.Rectangle(220, 420, 520, 70));

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

	public void slideLink(HSLFSlideShow ppt, HSLFSlide slide, ArrayList<String> rank) throws IOException {
		setBackground(slide, 2);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Link yang Paling Sering Digunakan");
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
		HSLFPictureData pd4123 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\linkFBv1.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew4123 = new HSLFPictureShape(pd4123);
		pictNew4123.setAnchor(new java.awt.Rectangle(120, 150, 700, 230));
		slide.addShape(pictNew4123);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();

		content.appendText(
				"Berdasarkan diagram diatas, terlihat bahwasanya link yang sering digunakan pada postingan adalah  ",
				false);
		content.appendText(rank.get(0) + " diikuti oleh " + rank.get(1) + " dan " + rank.get(2) + ".", false);

		content.setAnchor(new java.awt.Rectangle(100, 395, 760, 70));

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

	public void slideNetworkofPost(HSLFSlide slide) {
		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network of Post");
		title.setAnchor(new java.awt.Rectangle(500, 10, 450, 50));

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
		setTextColor(runTitle, version);

		slide.addShape(title);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();
		String jml_nodes = String.format("%s", "5221");
		String jml_edges = String.format("%s", "78342");
		String cluster1 = String.format("%s", "pro");
		String cluster2 = String.format("%s", "kontra");
		String cluster3 = String.format("%s", "netral");
		String cluster4 = String.format("%s", "netral1");
		String cluster5 = String.format("%s", "netral2");
		String jml_account = String.format("%s", "23423"); // banyak akun total
		// banyak akun tiap kelompok dalam persen
		String jml_account1 = String.format("%s", "58");
		String jml_account2 = String.format("%s", "58");
		String jml_account3 = String.format("%s", "58");
		String jml_account4 = String.format("%s", "58");
		String jml_account5 = String.format("%s", "58");
		String modularity1 = String.format("%s", "58");
		String modularity2 = String.format("%s", "58");
		String modularity3 = String.format("%s", "58");
		String modularity4 = String.format("%s", "58");
		String modularity5 = String.format("%s", "58");

		content.appendText("Network Postingan Facebook terdiri dari : ", false);
		content.appendText(
				"\tNodes\t\t: " + jml_nodes + "\n\tEdges\t\t: " + jml_edges + "\n\tBanyak Akun\t: " + jml_account,
				true);
		content.appendText("\n\nTerbentuk 5 kelompok dominan, yaitu : " + cluster1 + ", " + cluster2 + ", " + cluster3
				+ ", " + cluster4 + " dan " + cluster5 + ".", false);

		content.appendText(" ", true);
		content.appendText("- " + cluster1 + " [" + modularity1 + "% ]\n  Banyak Akun: " + jml_account1 + "%\n", true);
		content.appendText("- " + cluster2 + " [" + modularity2 + "% ]\n  Banyak Akun: " + jml_account2 + "%\n", true);
		content.appendText("- " + cluster3 + " [" + modularity3 + "% ]\n  Banyak Akun: " + jml_account3 + "%\n", true);
		content.appendText("- " + cluster4 + " [" + modularity4 + "% ]\n  Banyak Akun: " + jml_account4 + "%\n", true);
		content.appendText("- " + cluster5 + " [" + modularity5 + "% ]\n  Banyak Akun: " + jml_account5 + "%", true);

		content.setAnchor(new java.awt.Rectangle(500, 80, 430, 130));
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

	public void slideNetworkofPost2(HSLFSlide slide) {
		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network of Post");
		title.setAnchor(new java.awt.Rectangle(500, 10, 450, 50));

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
		setTextColor(runTitle, version);

		slide.addShape(title);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();

		String cluster1 = String.format("%s", "pro");
		String cluster2 = String.format("%s", "kontra");
		String cluster3 = String.format("%s", "netral");
		String cluster4 = String.format("%s", "netral1");
		String cluster5 = String.format("%s", "netral2");

		String hashtag1_cluster1 = String.format("%s", "#aku");
		String hashtag2_cluster1 = String.format("%s", "#2019presiden");
		String hashtag3_cluster1 = String.format("%s", "#deadpool");
		String hashtag1_cluster2 = String.format("%s", "#infinitywar");
		String hashtag2_cluster2 = String.format("%s", "#marvel");
		String hashtag3_cluster2 = String.format("%s", "#dukungaja");
		String hashtag1_cluster3 = String.format("%s", "#siapatakut");
		String hashtag2_cluster3 = String.format("%s", "#SNA");
		String hashtag3_cluster3 = String.format("%s", "#SMK");
		String hashtag1_cluster4 = String.format("%s", "#juara");
		String hashtag2_cluster4 = String.format("%s", "#IndonesiaIndicator");
		String hashtag3_cluster4 = String.format("%s", "#Badminton");
		String hashtag1_cluster5 = String.format("%s", "#Network");
		String hashtag2_cluster5 = String.format("%s", "#Ada Ular");
		String hashtag3_cluster5 = String.format("%s", "#Kebakaran");

		content.appendText("Hashtag yang sering digunakan oleh masing-masing kelompok: \n", false);
		content.appendText("\n- Hashtag " + cluster1 + " : \n  " + hashtag1_cluster1 + ", " + hashtag2_cluster1 + ", "
				+ hashtag3_cluster1, true);
		content.appendText("\n- Hashtag " + cluster2 + " : \n  " + hashtag1_cluster2 + ", " + hashtag2_cluster2 + ", "
				+ hashtag3_cluster2, true);
		content.appendText("\n- Hashtag " + cluster3 + " : \n  " + hashtag1_cluster3 + ", " + hashtag2_cluster3 + ", "
				+ hashtag3_cluster3, true);
		content.appendText("\n- Hashtag " + cluster4 + " : \n  " + hashtag1_cluster4 + ", " + hashtag2_cluster4 + ", "
				+ hashtag3_cluster4, true);
		content.appendText("\n- Hashtag " + cluster5 + " : \n  " + hashtag1_cluster5 + ", " + hashtag2_cluster5 + ", "
				+ hashtag3_cluster5, true);

		content.setAnchor(new java.awt.Rectangle(500, 80, 430, 130));
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

	public void slideNetworkofShare(HSLFSlide slide) {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network of Share");
		title.setAnchor(new java.awt.Rectangle(500, 10, 470, 50));

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
		setTextColor(runTitle, version);

		slide.addShape(title);

		// -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();

		String cluster1 = String.format("%s", "pro");
		String cluster2 = String.format("%s", "kontra");
		String cluster3 = String.format("%s", "netral");
		String cluster4 = String.format("%s", "netral1");
		String cluster5 = String.format("%s", "netral2");

		String akun1_cluster1_in = String.format("%s", "Spiderman");
		String akun2_cluster1_in = String.format("%s", "HawkEye");
		String akun3_cluster1_in = String.format("%s", "Black Widow");
		String akun1_cluster2_in = String.format("%s", "Black Panter");
		String akun2_cluster2_in = String.format("%s", "spedo");
		String akun3_cluster2_in = String.format("%s", "mastur");
		String akun1_cluster3_in = String.format("%s", "piyuu");
		String akun2_cluster3_in = String.format("%s", "benny");
		String akun3_cluster3_in = String.format("%s", "wenda");
		String akun1_cluster4_in = String.format("%s", "captain america");
		String akun2_cluster4_in = String.format("%s", "Hulk");
		String akun3_cluster4_in = String.format("%s", "ironman");
		String akun1_cluster5_in = String.format("%s", "Quick Silver");
		String akun2_cluster5_in = String.format("%s", "Superman");
		String akun3_cluster5_in = String.format("%s", "Batman");

		String modularity1 = String.format("%s", "58");
		String modularity2 = String.format("%s", "58");
		String modularity3 = String.format("%s", "58");
		String modularity4 = String.format("%s", "58");
		String modularity5 = String.format("%s", "58");

		content.appendText("Akun-akun yang dibagikan postingannya untuk masing-masing kelompok: ", false);
		content.appendText("\n- " + cluster1 + " (" + modularity1 + " %)\n  " + akun1_cluster1_in + "\n  "
				+ akun2_cluster1_in + "\n  " + akun3_cluster1_in, true);
		content.appendText("- " + cluster2 + " (" + modularity2 + " %)\n  " + akun1_cluster2_in + "\n  "
				+ akun2_cluster2_in + "\n  " + akun3_cluster2_in, true);
		content.appendText("- " + cluster3 + " (" + modularity3 + " %)\n  " + akun1_cluster3_in + "\n  "
				+ akun2_cluster3_in + "\n  " + akun3_cluster3_in, true);
		content.appendText("- " + cluster4 + " (" + modularity4 + " %)\n  " + akun1_cluster4_in + "\n  "
				+ akun2_cluster4_in + "\n  " + akun3_cluster4_in, true);
		content.appendText("- " + cluster5 + " (" + modularity5 + " %)\n  " + akun1_cluster5_in + "\n  "
				+ akun2_cluster5_in + "\n  " + akun3_cluster5_in, true);

		content.setAnchor(new java.awt.Rectangle(500, 78, 430, 130));
		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(100.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(16.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, version);

		slide.addShape(content);

		// =======content 2========//
		HSLFTextBox content2 = new HSLFTextBox();

		content2.setText("Ket: Nodes size berdasarkan in degree");

		content2.setAnchor(new java.awt.Rectangle(30, 470, 350, 385));

		HSLFTextParagraph titleP2 = content2.getTextParagraphs().get(0);
		titleP2.setAlignment(TextAlign.JUSTIFY);
		titleP2.setSpaceAfter(0.);
		titleP2.setSpaceBefore(0.);
		titleP2.setLineSpacing(110.0);

		HSLFTextRun runTitle2 = titleP2.getTextRuns().get(0);
		runTitle2.setFontSize(16.);
		runTitle2.setFontFamily("Calibri");
		runTitle2.setBold(false);
		setTextColor(runTitle2, version);

		slide.addShape(content2);

		// =======content Platform========//
		Platform(slide);

	}

	public void slideNetworkofShare2(HSLFSlide slide) {

		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network of Share");

		title.setAnchor(new java.awt.Rectangle(500, 10, 470, 50));

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
		setTextColor(runTitle, version);

		slide.addShape(title);
		//
		// // -------------------------Content------------------------
		HSLFTextBox content = new HSLFTextBox();
		String cluster1 = String.format("%s", "pro");
		String cluster2 = String.format("%s", "kontra");
		String cluster3 = String.format("%s", "netral");
		String cluster4 = String.format("%s", "netral1");
		String cluster5 = String.format("%s", "netral2");

		String akun1_cluster1_out = String.format("%s", "superman");
		String akun2_cluster1_out = String.format("%s", "batman");
		String akun3_cluster1_out = String.format("%s", "catwoman");
		String akun1_cluster2_out = String.format("%s", "deadpool");
		String akun2_cluster2_out = String.format("%s", "thanos");
		String akun3_cluster2_out = String.format("%s", "hulk");
		String akun1_cluster3_out = String.format("%s", "ironman");
		String akun2_cluster3_out = String.format("%s", "hackeye");
		String akun3_cluster3_out = String.format("%s", "spiderman");
		String akun1_cluster4_out = String.format("%s", "blackwidow");
		String akun2_cluster4_out = String.format("%s", "black panter");
		String akun3_cluster4_out = String.format("%s", "ant man");
		String akun1_cluster5_out = String.format("%s", "wolverine");
		String akun2_cluster5_out = String.format("%s", "Quick silver");
		String akun3_cluster5_out = String.format("%s", "thor");

		String modularity1 = String.format("%s", "58");
		String modularity2 = String.format("%s", "58");
		String modularity3 = String.format("%s", "58");
		String modularity4 = String.format("%s", "58");
		String modularity5 = String.format("%s", "58");
		//
		//
		content.appendText("Akun-akun yang sering membagikan postingan untuk masing masing kelompok: \n", false);
		content.appendText("\n- " + cluster1 + " (" + modularity1 + " %)\n  " + akun1_cluster1_out + "\n  "
				+ akun2_cluster1_out + "\n  " + akun3_cluster1_out, true);
		content.appendText("- " + cluster2 + " (" + modularity2 + " %)\n  " + akun1_cluster2_out + "\n  "
				+ akun2_cluster2_out + "\n  " + akun3_cluster2_out, true);
		content.appendText("- " + cluster3 + " (" + modularity3 + " %)\n  " + akun1_cluster3_out + "\n  "
				+ akun2_cluster3_out + "\n  " + akun3_cluster3_out, true);
		content.appendText("- " + cluster4 + " (" + modularity4 + " %)\n  " + akun1_cluster4_out + "\n  "
				+ akun2_cluster4_out + "\n  " + akun3_cluster4_out, true);
		content.appendText("- " + cluster5 + " (" + modularity5 + " %)\n  " + akun1_cluster5_out + "\n  "
				+ akun2_cluster5_out + "\n  " + akun3_cluster5_out, true);

		content.setAnchor(new java.awt.Rectangle(500, 78, 430, 130));
		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(0.);
		contentP.setSpaceBefore(0.);
		contentP.setLineSpacing(100.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(16.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(false);
		setTextColor(runContent, version);

		slide.addShape(content);

		// =======content 2========//
		HSLFTextBox content2 = new HSLFTextBox();

		content2.setText("Ket: Nodes size berdasarkan out degree");

		content2.setAnchor(new java.awt.Rectangle(30, 470, 350, 385));

		HSLFTextParagraph titleP2 = content2.getTextParagraphs().get(0);
		titleP2.setAlignment(TextAlign.JUSTIFY);
		titleP2.setSpaceAfter(0.);
		titleP2.setSpaceBefore(0.);
		titleP2.setLineSpacing(110.0);

		HSLFTextRun runTitle2 = titleP2.getTextRuns().get(0);
		runTitle2.setFontSize(16.);
		runTitle2.setFontFamily("Calibri");
		runTitle2.setBold(false);
		setTextColor(runTitle2, version);

		slide.addShape(content2);

		// =======content Platform========//

		Platform(slide);

	}

	public void slideConclution(HSLFSlide slide, String[] time1, String[] time2) {
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
