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

public class SNAInstagramAccount {

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

	public SNAInstagramAccount(String filename, int version, int width, int height, String tema) {
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

		// MasterTemplateIGAccount

		HSLFSlideShow ppt = new HSLFSlideShow();
		ppt.setPageSize(new java.awt.Dimension(width, height));

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\ig_logo.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(920, 0, 40, 40));

		// Slide 1
		HSLFSlide slide1 = ppt.createSlide();
		slideMaster(slide1);

		// ----- add picture-------------//
		HSLFPictureData pd41r = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\MasterTemplateIGAccount.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew41r = new HSLFPictureShape(pd41r);
		pictNew41r.setAnchor(new java.awt.Rectangle(40, 30, 300, 400));
		slide1.addShape(pictNew41r);

		HSLFPictureShape pictNew41r1 = new HSLFPictureShape(pd41r);
		pictNew41r1.setRotation(180);
		pictNew41r1.setAnchor(new java.awt.Rectangle(620, 110, 300, 400));
		slide1.addShape(pictNew41r1);

		// Slide 2
		HSLFSlide slide2 = ppt.createSlide();
		slideInformation(slide2);
		slide2.addShape(pictNew);

		// ------------- add picture -----------//
		HSLFPictureData pd3 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\wordcloudInstagramAccount.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3 = new HSLFPictureShape(pd3);
		pictNew3.setAnchor(new java.awt.Rectangle(630, 80, 250, 250));
		slide2.addShape(pictNew3);

		// Slide 3
		HSLFSlide slide3 = ppt.createSlide();
		slideWordcloud(slide3);
		slide3.addShape(pictNew);
		// ------------- add picture -----------//
		HSLFPictureData pd31t3q = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\wordcloudInstagramAccount.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew31t3q = new HSLFPictureShape(pd31t3q);
		pictNew31t3q.setAnchor(new java.awt.Rectangle(590, 80, 350, 350));
		slide3.addShape(pictNew31t3q);

		// Slide 4
		HSLFSlide slide4 = ppt.createSlide();
		slidePostTypes(slide4);
		slide4.addShape(pictNew);

		// ------------- add picture -----------//
		HSLFPictureData pd1t3x = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\donatInstagramAccount.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1t3x = new HSLFPictureShape(pd1t3x);
		pictNew1t3x.setAnchor(new java.awt.Rectangle(100, 100, 400, 340));
		slide4.addShape(pictNew1t3x);

		// Slide 5
		HSLFSlide slide5 = ppt.createSlide();
		slideTopHashtag(slide5);
		slide5.addShape(pictNew);
		
		// ------------- add picture -----------//
		HSLFPictureData pd1 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barHashtagTeleChannel.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
		pictNew1.setAnchor(new java.awt.Rectangle(220, 70, 600, 320));
		slide5.addShape(pictNew1);
		

		// Slide 6
		HSLFSlide slide6 = ppt.createSlide();
		slideTopMention(slide6);
		slide6.addShape(pictNew);
		
		// ------------- add picture -----------//
		HSLFPictureData pd11 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barHashtagTeleChannel.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew11 = new HSLFPictureShape(pd11);
		pictNew11.setAnchor(new java.awt.Rectangle(220, 70, 600, 320));
		slide6.addShape(pictNew11);
		
		

		// Slide 7
		HSLFSlide slide7 = ppt.createSlide();
		slideTopPostbyLikes(slide7);
		slide7.addShape(pictNew);
		
		// ------------- add picture -----------//
		HSLFPictureData pd12 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barHashtagTeleChannel.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew12 = new HSLFPictureShape(pd12);
		pictNew12.setAnchor(new java.awt.Rectangle(220, 70, 600, 320));
		slide7.addShape(pictNew12);
		

		// Slide 8
		HSLFSlide slide8 = ppt.createSlide();
		slideTopPostbyComment(slide8);
		slide8.addShape(pictNew);

		// ------------- add picture -----------//
		HSLFPictureData pd13 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barHashtagTeleChannel.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew13 = new HSLFPictureShape(pd13);
		pictNew13.setAnchor(new java.awt.Rectangle(220, 70, 600, 320));
		slide8.addShape(pictNew13);
		
		
		// Slide 9
		HSLFSlide slide9 = ppt.createSlide();
		slideTopTaggedAccount(slide9);
		slide9.addShape(pictNew);

		// ------------- add picture -----------//
		HSLFPictureData pd14 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barHashtagTeleChannel.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew14 = new HSLFPictureShape(pd14);
		pictNew14.setAnchor(new java.awt.Rectangle(220, 70, 600, 320));
		slide9.addShape(pictNew14);
		
		// Slide 10
		HSLFSlide slide10 = ppt.createSlide();
		slideActivityGraph(slide10);
		slide10.addShape(pictNew);

		// ------------- add picture -----------//
		HSLFPictureData pd15 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barHashtagTeleChannel.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew15 = new HSLFPictureShape(pd15);
		pictNew15.setAnchor(new java.awt.Rectangle(220, 85, 600, 300));
		slide10.addShape(pictNew15);
		
		
		// Slide 11
		HSLFSlide slide11 = ppt.createSlide();
		slideTopMentionAccount(slide11);
		slide11.addShape(pictNew);

		// ------------- add picture -----------//
		HSLFPictureData pd1a = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\barHashtagTeleChannel.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1a = new HSLFPictureShape(pd1a);
		pictNew1a.setAnchor(new java.awt.Rectangle(220, 70, 600, 320));
		slide11.addShape(pictNew1a);
		
		// Slide 12
		HSLFSlide slide12 = ppt.createSlide();
		slideClosenessAnalysis(slide12);
		slide12.addShape(pictNew);

		// Slide 13
		HSLFSlide slide13 = ppt.createSlide();
		slideGeneralNetwork(slide13);
		slide13.addShape(pictNew);

		// Slide 14
		HSLFSlide slide14 = ppt.createSlide();
		slideMutualEdges(slide14);
		slide14.addShape(pictNew);

		// Slide 15
		HSLFSlide slide15 = ppt.createSlide();
		slideNetworkbyPage(slide15);
		slide15.addShape(pictNew);

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

		else if (ver == 3) {
			HSLFShape shape = new HSLFAutoShape(ShapeType.RECT);
			shape.setAnchor(new java.awt.Rectangle(0, 0, width, height));
			HSLFFill fill = shape.getFill();
			fill.setBackgroundColor(new Color(244, 241, 225));
			fill.setForegroundColor(new Color(244, 241, 225));
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
		title.setRotation(270);
		title.setAnchor(new java.awt.Rectangle(7, 0, 50, 540));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.CENTER);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(36.);
		runTitle.setBold(true);
		runTitle.setFontFamily("SimSun-ExtB (Headings)");

		if (versi == version) {
			runTitle.setFontColor(Color.BLACK);
		} else {
			runTitle.setFontColor(Color.WHITE);
		}

		slide.addShape(title);

		// =========garis===========//
		Garis(slide, 70, 14, 1, 511, versi);

	}

	public void Garis(HSLFSlide slide, int a, int b, int c, int d, int versi) {
		HSLFTextBox content = new HSLFTextBox();
		if (versi == 1) {
			content.setLineColor(Color.GRAY);

		} else {
			content.setLineColor(new Color(25, 27, 14));
		}

		content.setLineWidth(30);
		content.setRotation(180);
		content.setAnchor(new java.awt.Rectangle(a, b, c, d));
		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		slide.addShape(content);
	}

	public void slideMaster(HSLFSlide slide) {
		setBackground(slide, version);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("SOCIAL NETWORK ANALYSIS");
		title.setAnchor(new java.awt.Rectangle(10, 220, 950, 50));

		HSLFTextParagraph titleP = title.getTextParagraphs().get(0);
		titleP.setAlignment(TextAlign.CENTER);
		titleP.setSpaceAfter(0.);
		titleP.setSpaceBefore(0.);

		HSLFTextRun runTitle = titleP.getTextRuns().get(0);
		runTitle.setFontSize(54.);
		runTitle.setFontFamily("SimSun-ExtB");
		runTitle.setBold(true);
		runTitle.setUnderlined(false);
		runTitle.setFontColor(Color.BLACK);

		slide.addShape(title);

		// --------------------Investigation--------------------
		String nama_akun = "@Ridwan";

		HSLFTextBox issue = new HSLFTextBox();
		issue.setText("INVESTIGASI AKUN " + nama_akun);
		issue.setAnchor(new java.awt.Rectangle(10, 280, 950, 50));
		HSLFTextParagraph issueP = issue.getTextParagraphs().get(0);
		issueP.setAlignment(TextAlign.CENTER);
		HSLFTextRun runIssue = issueP.getTextRuns().get(0);
		runIssue.setFontSize(23.);
		runIssue.setFontFamily("SimSun-ExtB");
		setTextColor(runIssue, version);
		slide.addShape(issue);

		// -------------------------Date------------------------
		HSLFTextBox date = new HSLFTextBox();
		Date now = new Date(System.currentTimeMillis());
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy");
		date.setText(dateFormat.format(now));
		date.setAnchor(new java.awt.Rectangle(10, 315, 950, 50));
		HSLFTextParagraph dateP = date.getTextParagraphs().get(0);
		dateP.setAlignment(TextAlign.CENTER);
		HSLFTextRun runDate = dateP.getTextRuns().get(0);
		runDate.setFontSize(18.);
		runDate.setFontFamily("SimSun-ExtB");
		setTextColor(runDate, version);
		slide.addShape(date);
	}

	public void slideInformation(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Information", 3);

		// -------------------------Content------------------------//
		HSLFTextBox content = new HSLFTextBox();
		content.setText("Basic Data");
		content.appendText("\n\n\n\n\n\n\nBasic Network Info", false);
		content.appendText("\n\n\n\n\n\n\nBasic Statistic Info", false);
		content.setAnchor(new java.awt.Rectangle(140, 60, 600, 480));

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.JUSTIFY);
		contentP.setSpaceAfter(2.);
		contentP.setSpaceBefore(10.);
		contentP.setLineSpacing(110.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(15.);
		runContent.setFontFamily("SimSun-ExtB (Body)");
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
		String following = "Ridwan Kamil";
		String follower = "Ridwan Kamil";
		String post = "Ridwan Kamil";
		String average_like = "Ridwan Kamil";
		String user_ID = "Ridwan Kamil";
		String average_comment = "Ridwan Kamil";

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Display Name\t:  " + display_name);
		content1.appendText("\nUsername\t\t:  " + username, false);
		content1.appendText("\nBiography\t\t:  " + bio, false);
		content1.appendText("\nURL\t\t:  " + url, false);
		content1.appendText("\nLocation\t\t:  " + location, false);
		content1.appendText("\n\n\nFollowings\t:  " + following + " accounts", false);
		content1.appendText("\nFollowers\t\t:  " + follower + " accounts", false);
		content1.appendText("\nPosts\t\t:  " + post + " posts", false);
		content1.appendText("\nUser ID\t\t:  " + user_ID, false);
		content1.appendText("\n\n\n\nAverage likes per post\t\t:  " + average_like + " likes", false);
		content1.appendText("\nAverage comment per post\t\t:  " + average_comment + " comments", false);
		content1.appendText("\nFollower active based on like\t:  " + follower_active + " %", false);
		content1.setAnchor(new java.awt.Rectangle(140, 80, 600, 480));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.JUSTIFY);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(15.);
		runContent1.setFontFamily("SimSun-ExtB (Body)");
		setTextColor(runContent1, version);
		slide.addShape(content1);
	}

	public void slideWordcloud(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Wordcloud", 3);
		// -------------------------Content2------------------------//

		String username = "Ridwan Kamil";
		String jml_data = "40";
		String key1 = "Ridwan Kamil";
		String key2 = "Ridwan Kamil";
		String key3 = "Ridwan Kamil";
		String key4 = "Ridwan Kamil";
		String key5 = "Ridwan Kamil";

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Dengan menggunakan " + jml_data + " data kiriman terakhir dari akun " + username);
		content1.appendText(
				", terbentuk visualisasi caption yang menunjukkan tema kiriman yang sering diunggah oleh akun tersebut.",
				false);
		content1.appendText(" Ukuran font menunjukkan frekuensi kemunculan kata dalam konten caption.", false);
		content1.appendText(
				"\n\nWordcloud di samping menunjukkan bahwa terdapat beberapa kata yang dominan digunakan oleh akun "
						+ username,
				false);
		content1.appendText(", yakni " + key1 + ", " + key2 + ", " + key3 + ", " + key4 + ", dan " + key5 + ".", false);
		content1.setAnchor(new java.awt.Rectangle(140, 80, 420, 480));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.JUSTIFY);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(19.);
		runContent1.setFontFamily("SimSun-ExtB (Body)");
		setTextColor(runContent1, version);
		slide.addShape(content1);

	}

	public void slidePostTypes(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "PostTypes", 3);
		// -------------------------Content2------------------------//

		String username = "Ridwan Kamil";
		String post_type = "photo";
		String jml_post_type = "500";

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Berdasarkan hasil visualisasi di samping, diketahui bahwa mayoritas post dari akun "
				+ username + " dikirim dalam bentuk " + post_type);
		content1.appendText(
				". Sebanyak " + jml_post_type + " post diupload oleh akun tersebut dalam bentuk " + post_type + "",
				false);

		content1.setAnchor(new java.awt.Rectangle(530, 150, 370, 400));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.JUSTIFY);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(110.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(20.);
		runContent1.setFontFamily("SimSun-ExtB (Body)");
		setTextColor(runContent1, version);
		slide.addShape(content1);

	}

	public void slideTopHashtag(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Top Hashtag", 3);
		// -------------------------Content2------------------------//

		String username = "Ridwan Kamil";
		String hashtag_terbesar = "#kopasus";
		String jml_hashtag_terbesar = "200";

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Berikut adalah 5 hashtag yang paling sering digunakan di caption akun " + username);
		content1.appendText(". Hashtag " + hashtag_terbesar + " digunakan sebanyak " + jml_hashtag_terbesar + " kali.",
				false);

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

	public void slideTopMention(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Top Mention", 3);
		// -------------------------Content2------------------------//

		String username = "Ridwan Kamil";
		String akun_terbesar = "#kopasus";

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Berikut adalah 5 hashtag yang paling sering disebut oleh akun " + username);
		content1.appendText(". Akun " + akun_terbesar
				+ " menjadi akun yang paling sering dimention menunjukkan bahwa akun tersebut memiliki hubungan dan pengaruh yang kuat terhadap akun."
				+ username, false);

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

	public void slideTopPostbyLikes(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Top Post by Likes", 3);
		// -------------------------Content2------------------------//

		String username = "Ridwan Kamil";
		String akun_terbesar = "#kopasus";

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Gambar di atas menunjukkan 5 kiriman yang paling banyak disukai dari akun " + username);

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

	public void slideTopPostbyComment(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Top Post by Comment", 3);
		// -------------------------Content2------------------------//

		String username = "Ridwan Kamil";

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Berikut adalah 5 post yang paling banyak mendapat komentar di akun " + username);

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

	public void slideTopTaggedAccount(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Top Tagged Account", 3);
		// -------------------------Content2------------------------//

		String username = "Ridwan Kamil";
		String influencer = "@bts.bighitofficial";

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Grafik berikut menggambarkan 5 akun yang paling sering di-tag oleh akun " + username);

		content1.appendText(". Akun yang paling sering di-tag oleh akun " + username + " adalah " + influencer
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

	public void slideActivityGraph(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Activity Graph", 3);
		// -------------------------Content2------------------------//

		String username = "Ridwan Kamil";
		String influencer = "@bts.bighitofficial";
		String tanggal_awal = "1/2/2018";
		String tanggal_akhir = "7/2/2018";
		String tanggal_puncak = "3/2/108";
		String jml_tweet_puncak = "9090";
		String jml_data = "99";
		String jml_akun = "44";

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText(
				"Grafik berikut menggambarkan aktifitas akun " + username + " selama satu tahun kebelakang.\n\n\n\n");
		content1.appendText(
				"\n\n\n\n\n\n\n\n•Puncak pembicaraan tertinggi dari issue " + tema + " terjadi pada tanggal "
						+ tanggal_puncak + " dengan jumlah pembicaraan mencapai " + jml_tweet_puncak + " post/hours.",
				false);
		content1.appendText("\n•Jumlah pembicaraan issue ini dalam kurun " + tanggal_awal + " sampai tanggal "
				+ tanggal_akhir + " sebanyak " + jml_data
				+ " Posts, dengan jumlah massa facebook yang membicarakan sebanyak " + jml_akun + " akun.", false);

		content1.setAnchor(new java.awt.Rectangle(105, 35, 775, 120));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.CENTER);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(105.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(19.);
		runContent1.setFontFamily("SimSun-ExtB (Body)");
		setTextColor(runContent1, version);
		slide.addShape(content1);

	}

	public void slideClosenessAnalysis(HSLFSlide slide) {
		setBackground(slide, version);
		judul(slide, "Top Closeness Analysis", 3);
		// -------------------------Content2------------------------//

		String username = "Ridwan Kamil";
		String jml_data = "99";
		String akun1 = "44";
		String akun2 = "44";
		String akun3 = "44";

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Tabel-tabel di bawah menunjukkan kedekatan antara akun " + username
				+ " terhadap akun-akun lain, dengan menggunakan count(data) post terakhir" + jml_data + ".\n\n\n\n");

		content1.appendText("\nAkun yang paling sering menyukai kiriman dari  akun " + username + ", adalah: " + akun1
				+ ", " + akun2 + ", " + akun3 + ".", false);

		content1.appendText("\n\nSedangkan akun yang paling sering mengomentari kiriman akun " + username + ", adalah: "
				+ akun1 + ", " + akun2 + ", " + akun3 + ".", false);

		content1.setAnchor(new java.awt.Rectangle(105, 40, 775, 120));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.CENTER);
		contentP1.setSpaceAfter(2.);
		contentP1.setSpaceBefore(10.);
		contentP1.setLineSpacing(105.0);

		HSLFTextRun runContent1 = contentP1.getTextRuns().get(0);
		runContent1.setFontSize(19.);
		runContent1.setFontFamily("SimSun-ExtB (Body)");
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

	public void slideMutualEdges(HSLFSlide slide) {
		setBackground(slide, 1);
		judul(slide, "Mutual Edges Network", 1);
		// -------------------------Content2------------------------//

		String username = "Ridwan Kamil";
		String post_type = "photo";
		String jml_post_type = "500";

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Network mutual edges berfungsi untuk menunjukkan akun-akun yang saling follow dengan akun  "
				+ username);

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

	public void slideNetworkbyPage(HSLFSlide slide) {
		setBackground(slide, 1);
		judul(slide, "Network by Page Rank", 1);
		// -------------------------Content2------------------------//

		String username = "Ridwan Kamil";

		HSLFTextBox content1 = new HSLFTextBox();
		content1.setText("Akun " + username + " adalah akun yang paling banyak dirujuk oleh akun-akun besar.");

		content1.setAnchor(new java.awt.Rectangle(120, 400, 760, 100));

		HSLFTextParagraph contentP1 = content1.getTextParagraphs().get(0);
		contentP1.setAlignment(TextAlign.CENTER);
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
