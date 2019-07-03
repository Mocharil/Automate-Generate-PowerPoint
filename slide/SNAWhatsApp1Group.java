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

public class SNAWhatsApp1Group {

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

	public SNAWhatsApp1Group(String filename, int version, int width, int height, String tema) {
		this.filename = filename;
		this.version = version;
		this.height = height;
		this.width = width;
		this.tema = tema;
	}

	public void execute() throws Exception {

		// -----------------MongoDB--------------------//
		// final MongoCredential credential =
		// MongoCredential.createScramSha1Credential(mongoUsername, mongoSource,
		// mongoPassword.toCharArray());
		// ServerAddress serverAddress = new ServerAddress(mongoHost,
		// Integer.parseInt(mongoPort));
		// MongoClient mongoClient = new MongoClient(serverAddress, new
		// ArrayList<MongoCredential>() {
		// /**
		// *
		// */
		// private static final long serialVersionUID = 1L;
		//
		// {
		// add(credential);
		// }
		// });
		//
		// MongoDatabase database = mongoClient.getDatabase("criteria_twitter");
		// MongoCollection<Document> collection =
		// database.getCollection("330a59fae963a979e2d96a9ec91e8b7c");
		//
		// MongoCursor<Document> cursor = collection.find().iterator();
		// ArrayList<Document> list = new ArrayList<Document>();
		//
		// try {
		// while (cursor.hasNext()) {
		// list.add(cursor.next());
		//
		// }
		// } catch (Exception e) {
		// System.out.println("gagal");
		// } finally {
		// cursor.close();
		// }


		HSLFSlideShow ppt = new HSLFSlideShow();
		ppt.setPageSize(new java.awt.Dimension(width, height));

		// Slide 1
		HSLFSlide slide1 = ppt.createSlide();
		slideMaster(slide1);

		// Slide 2
		HSLFSlide slide21 = ppt.createSlide();
		slideSingleGroup(slide21);

		// Slide 3
		HSLFSlide slide2 = ppt.createSlide();
		slideInformation(slide2);

		// ------------- add picture -----------//
		HSLFPictureData pd1t3x = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\donatInstagramAccount.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1t3x = new HSLFPictureShape(pd1t3x);
		pictNew1t3x.setAnchor(new java.awt.Rectangle(80, 100, 400, 340));
		slide2.addShape(pictNew1t3x);

		// Slide 4
		HSLFSlide slide4 = ppt.createSlide();
		slideRawStatistics(slide4);

		// Slide 5
		HSLFSlide slide3 = ppt.createSlide();
		slideWordcloud(slide3);

		// ------------- add picture -----------//
		HSLFPictureData pd3 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\wordcloudInstagramAccount.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3 = new HSLFPictureShape(pd3);
		pictNew3.setAnchor(new java.awt.Rectangle(80, 100, 400, 340));
		slide3.addShape(pictNew3);

		// Slide 6
		HSLFSlide slide5 = ppt.createSlide();
		slideMostResponseMessage(slide5);

		// ------------- add picture -----------//
		HSLFPictureData pd1 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barHashtagTeleChannel.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
		pictNew1.setAnchor(new java.awt.Rectangle(220, 150, 600, 320));
		slide5.addShape(pictNew1);

		// Slide 6
		HSLFSlide slide6 = ppt.createSlide();
		slideMostResponseNumber(slide6);

		// ------------- add picture -----------//
		HSLFPictureData pd11 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barHashtagTeleChannel.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew11 = new HSLFPictureShape(pd11);
		pictNew11.setAnchor(new java.awt.Rectangle(220, 150, 600, 320));
		slide6.addShape(pictNew11);

		// Slide 7
		HSLFSlide slide7 = ppt.createSlide();
		slideMostChattyNumber(slide7);

		// ------------- add picture -----------//
		HSLFPictureData pd12 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barHashtagTeleChannel.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew12 = new HSLFPictureShape(pd12);
		pictNew12.setAnchor(new java.awt.Rectangle(220, 150, 600, 320));
		slide7.addShape(pictNew12);

		// Slide 8
		HSLFSlide slide8 = ppt.createSlide();
		slideMostSharedLink(slide8);

		// ------------- add picture -----------//
		HSLFPictureData pd13 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barHashtagTeleChannel.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew13 = new HSLFPictureShape(pd13);
		pictNew13.setAnchor(new java.awt.Rectangle(220, 150, 600, 320));
		slide8.addShape(pictNew13);

		// Slide 9
		HSLFSlide slide9 = ppt.createSlide();
		slideGeographic(slide9);

		// ------------- add picture -----------//
		HSLFPictureData pd14 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barHashtagTeleChannel.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew14 = new HSLFPictureShape(pd14);
		pictNew14.setAnchor(new java.awt.Rectangle(220, 150, 600, 320));
		slide9.addShape(pictNew14);

		// Slide 13
		HSLFSlide slide13 = ppt.createSlide();
		slideGeneralNetwork(slide13);

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

		else if (ver == 2) {
			HSLFShape shape = new HSLFAutoShape(ShapeType.RECT);
			shape.setAnchor(new java.awt.Rectangle(0, 0, width, height));
			HSLFFill fill = shape.getFill();
			fill.setBackgroundColor(new Color(84, 130, 53));
			fill.setForegroundColor(new Color(84, 130, 53));
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
		title.setAnchor(new java.awt.Rectangle(10, 20, 940, 100));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.RIGHT);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(40.);
		runTitle.setItalic(true);
		runTitle.setFontFamily("Century Schoolbook");

		if (versi == version) {
			runTitle.setFontColor(Color.BLACK);
		} else {
			runTitle.setFontColor(Color.WHITE);
		}

		slide.addShape(title);

	}

	public void slideMaster(HSLFSlide slide) {
		setBackground(slide, 2);
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
		runTitle.setFontSize(60.);
		runTitle.setFontFamily("Times New Roman");
		runTitle.setBold(true);
		runTitle.setItalic(true);
		runTitle.setUnderlined(false);
		runTitle.setFontColor(Color.red);

		HSLFTextParagraph titleP1 = title.getTextParagraphs().get(2);
		HSLFTextRun runTitle1 = titleP1.getTextRuns().get(0);
		runTitle1.setFontSize(60.);
		runTitle1.setFontFamily("Times New Roman");
		runTitle1.setBold(true);
		runTitle1.setItalic(true);
		runTitle1.setUnderlined(false);
		setTextColor(runTitle1, version);
		slide.addShape(title);

		// --------------------Investigation--------------------
		HSLFTextBox issue = new HSLFTextBox();
		String nama_group = String.format("%s", "NAMA GROUP");
		issue.appendText("INVESTIGASI Group Whatsapp " + nama_group, false);
		issue.setAnchor(new java.awt.Rectangle(60, 330, 900, 50));
		HSLFTextParagraph issueP = issue.getTextParagraphs().get(0);
		HSLFTextRun runIssue = issueP.getTextRuns().get(0);
		runIssue.setFontSize(24.);
		runIssue.setFontFamily("Century Schoolbook");
		runIssue.setBold(true);
		runIssue.setItalic(true);
		setTextColor(runIssue, version);
		slide.addShape(issue);

		// -------------------------Date------------------------
		HSLFTextBox date = new HSLFTextBox();
		Date now = new Date(System.currentTimeMillis());
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
		date.setText(dateFormat.format(now));
		date.setAnchor(new java.awt.Rectangle(60, 430, 200, 50));
		HSLFTextParagraph dateP = date.getTextParagraphs().get(0);
		dateP.setAlignment(TextAlign.LEFT);
		HSLFTextRun runDate = dateP.getTextRuns().get(0);
		runDate.setFontSize(12.);
		runDate.setFontFamily("Century Schoolbook");
		setTextColor(runDate, version);
		slide.addShape(date);
	}

	public void slideSingleGroup(HSLFSlide slide) {
		setBackground(slide, 3);

		// -------------------------Content------------------------//
		HSLFTextBox content = new HSLFTextBox();
		content.setText("SINGLE GROUP");

		content.setAnchor(new java.awt.Rectangle(50, 330, 600, 50));

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.LEFT);
		contentP.setSpaceAfter(2.);
		contentP.setSpaceBefore(10.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(60.);
		runContent.setFontFamily("Calibri Light (Headings)");
		setTextColor(runContent, version);
		slide.addShape(content);
	}

	public void slideInformation(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Information", 3);

		// -------------------------Content------------------------//
		HSLFTextBox content = new HSLFTextBox();
		content.setText("Basic Data");
		content.setAnchor(new java.awt.Rectangle(490, 90, 430, 410));
		content.setLineColor(Color.BLACK);
		content.setLineWidth(2);

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(2.);
		contentP.setSpaceBefore(10.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(15.);
		runContent.setFontFamily("Calibri");
		runContent.setBold(true);
		runContent.setUnderlined(true);
		setTextColor(runContent, version);
		slide.addShape(content);

		// -------------------------Content2------------------------//
		String follower_active = "Ridwan Kamil";
		String display_name = "Ridwan Kamil";
		String username = "Ridwan Kamil";
		String bio = "Ridwan Kamil";
		String url = "Ridwan Kamil";
		String location = "Ridwan Kamil";

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Display Group Name\t:  " + display_name);
		content1.appendText("\nNumber of Member\t:  " + username, false);
		content1.appendText("\nAdmin's Number\t:  " + bio, false);
		content1.appendText("\nCreated at\t\t:  " + url, false);
		content1.appendText("\n\n\nBrief Description about the group :  " + location, false);
		content1.setAnchor(new java.awt.Rectangle(490, 110, 430, 390));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.JUSTIFY);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(16.);
		runContent1.setFontFamily("Calibri");
		setTextColor(runContent1, version);
		slide.addShape(content1);

	}

	public void slideWordcloud(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Wordcloud", 3);
		// -------------------------Content2------------------------//

		String nama_group = "Ridwan Kamil";
		String waktu = "photo";
		String jml_percakapan = "500";
		String key1 = "as";
		String key2 = "as";
		String key3 = "as";

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText(
				"Wordcloud di samping menunjukkan perbincangan akhir-akhir ini yang dibahas oleh akun " + nama_group);
		content1.appendText(
				", terdata ada " + jml_percakapan + " percakapan selama tanggal " + waktu + " hingga " + waktu, false);

		content1.appendText("\n\nHasil visualisasi disamping menunjukkan kata kata yang diperbincangkan, ", false);
		content1.appendText(
				"terlihat kata yang paling banyak disebutkan direpresentasikan dengan ukuran tulisan yang semakin besar yaitu "
						+ key1 + ", " + key2 + ", dan " + key3 + ".",
				false);

		content1.setAnchor(new java.awt.Rectangle(530, 120, 380, 400));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.JUSTIFY);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(20.);
		runContent1.setFontFamily("Calibri");
		setTextColor(runContent1, version);
		slide.addShape(content1);

	}

	public void slideRawStatistics(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Raw Statistics", 3);

	}

	public void slideMostResponseMessage(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Most Response Message", 3);
		// -------------------------Content2------------------------//

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Berikut adalah 3 message yang paling banyak di reply");

		content1.setAnchor(new java.awt.Rectangle(50, 80, 760, 100));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.LEFT);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(20.);
		runContent1.setFontFamily("Calibri");
		setTextColor(runContent1, version);
		slide.addShape(content1);

	}

	public void slideMostResponseNumber(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Most Response Number", 3);
		// -------------------------Content2------------------------//

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Berikut adalah 4 nomor hp yang messages nya direspon oleh rekan grup");

		content1.setAnchor(new java.awt.Rectangle(50, 80, 760, 100));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.LEFT);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(20.);
		runContent1.setFontFamily("Calibri");
		setTextColor(runContent1, version);
		slide.addShape(content1);
	}

	public void slideMostChattyNumber(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Most Chatty Number", 3);
		// -------------------------Content2------------------------//

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Berikut adalah 4 nomor hp yang messages nya direspon oleh rekan grup");

		content1.setAnchor(new java.awt.Rectangle(50, 80, 760, 100));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.LEFT);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(20.);
		runContent1.setFontFamily("Calibri");
		setTextColor(runContent1, version);
		slide.addShape(content1);

	}

	public void slideMostSharedLink(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Most Shared Link", 3);
		// -------------------------Content2------------------------//

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Berikut adalah 4 nomor domain yang paling banyak dibagikan");

		content1.setAnchor(new java.awt.Rectangle(50, 80, 760, 100));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.LEFT);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(20.);
		runContent1.setFontFamily("Calibri");
		setTextColor(runContent1, version);
		slide.addShape(content1);

	}

	public void slideGeographic(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Geographic Phone Number Region", 3);
		// -------------------------Content2------------------------//

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Berikut adalah representasi dari asal nomor hp");

		content1.setAnchor(new java.awt.Rectangle(50, 80, 760, 100));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.LEFT);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(20.);
		runContent1.setFontFamily("Calibri");
		setTextColor(runContent1, version);
		slide.addShape(content1);

	}


	public void slideGeneralNetwork(HSLFSlide slide) {
		setBackground(slide, 1);
		judul(slide, "General Network", 1);
		// -------------------------Content2------------------------//

		String username = "Ridwan Kamil";
		String post_type = "photo";
		String jml_post_type = "500";

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Network di samping menunjukkan posisi akun " + username
				+ " dalam jaringan, dengan menggunakan data following of following dari kedua akun tersebut.");

		content1.setAnchor(new java.awt.Rectangle(530, 150, 370, 400));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.JUSTIFY);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(20.);
		runContent1.setFontFamily("SimSun-ExtB (Body)");
		setTextColor(runContent1, 1);
		slide.addShape(content1);

	}


	public void slideConclusion(HSLFSlide slide) {
		judul(slide, "Kesimpulan", 3);
	}
}
