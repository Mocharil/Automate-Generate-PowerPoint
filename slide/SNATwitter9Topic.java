package com.ebdesk.report.slide;

import java.awt.Color;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.Scanner;
import java.util.List;
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
import java.util.Date;

public class SNATwitter9Topic {

	private int version;
	private String filename;
	private int width;
	private int height;
	private String tema;
	private static final char DEFAULT_SEPARATOR = ',';
	private static final char DEFAULT_QUOTE = '"';

	public SNATwitter9Topic(String filename, int version, int width, int height, String tema) {
		this.filename = filename;
		this.version = version;
		this.height = height;
		this.width = width;
		this.tema = tema;

	}

	public void execute() throws Exception {

		// =============TABEL============//
		List<HSLFTable> tabelbyfollower = new ArrayList<HSLFTable>();
		List<HSLFTable> tabelbyreach = new ArrayList<HSLFTable>();
		List<HSLFTable> tabeltweet = new ArrayList<HSLFTable>();
		for (int i = 0; i < 9; i++) {
			tabelbyfollower.add(tablev2(
					"C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_top10accountbyfollower.csv", 11,
					5));
			tabelbyreach.add(tablev2(
					"C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_top10accountbyreach.csv", 11, 5));
			tabeltweet.add(
					tablev2("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_top10tweet.csv", 6, 7));
		}

		HSLFTable table_hashtag_growth_kon = tablev2(
				"C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_top10hastaggrowth.csv", 11, 7);
		
		
		String[] tokoh = { "Jokowi", "Prabowo", "Anies Baswedan", "Jusuf Kalla", "AHY", "Gatot Nurmantyo", "Mahfud MD",
				"Cakimin", "Airlangga Hartanto" };

		// ====================CSV==================//
		String csvFile = "C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\info.csv";
		Scanner scanner = new Scanner(new File(csvFile));
		Map<String, String> map = new HashMap<String, String>();
		while (scanner.hasNext()) {
			List<String> line = parseLine(scanner.nextLine());
			map.put(line.get(0), line.get(1));

		}
		scanner.close();

		// ================DATA================//
		String[] time = new String[5];
		time = map.get("mindate").split("\\D");
		DateFormat df4 = new SimpleDateFormat("dd MM yyyy");
		DateFormat df = new SimpleDateFormat("dd MMMM yyyy");
		String waktu = new String();
		Date d1 = new Date();

		waktu = time[2] + " " + time[1] + " " + time[0];
		d1 = df4.parse(waktu);
		String time1 = df.format(d1); // tanggal awal

		time = map.get("maxdate").split("\\D");
		waktu = time[2] + " " + time[1] + " " + time[0];
		d1 = df4.parse(waktu);
		String time2 = df.format(d1);// tanggal akhir
		time = map.get("access_date").split("\\D");
		waktu = time[2] + " " + time[1] + " " + time[0];
		d1 = df4.parse(waktu);
		String akses = df.format(d1); // tanggal akses
		String tanggal_comparison_awal = map.get("awal_komparasi_pro"); // tanggal comparison
		String tanggal_comparison_akhir = map.get("akhir_komparasi_pro"); // tanggal comparison

		// -------------INFO------------//
		String Akun1 = "@" + map.get("tweet1acc");
		String Akun2 = "@" + map.get("tweet2acc");
		String Akun3 = "@" + map.get("tweet3acc");
		String Tweetakun1 = "\"" + map.get("tweet1") + "\"";
		String Tweetakun2 = "\"" + map.get("tweet2") + "\"";
		String Tweetakun3 = "\"" + map.get("tweet3") + "\"";
		String Url_image_akun1 = map.get("tweet1accimg").replace("normal", "200x200");
		String Url_image_akun2 = map.get("tweet2accimg").replace("normal", "200x200");
		String Url_image_akun3 = map.get("tweet3accimg").replace("normal", "200x200");
		String jumlah_data = map.get("tweet_count");

		// ----------------Tweet Trend--------------//
		String mean_pro = map.get("mean_pro");
		if (map.get("mean_pro").length() > 8) {
			mean_pro = map.get("mean_pro").substring(0, 8);
		}
		String q1_pro = map.get("q1_pro");
		String q2_pro = map.get("q2_pro");
		String q3_pro = map.get("q3_pro");
		String Total_pro = map.get("total_pro");

		String mean_kon = map.get("mean_kontra");
		if (map.get("mean_kontra").length() > 8) {
			mean_kon = map.get("mean_kontra").substring(0, 8);
		}

		String q1_kon = map.get("q1_kontra");
		String q2_kon = map.get("q2_kontra");
		String q3_kon = map.get("q3_kontra");
		String Total_kon = map.get("total_kontra");

		// ---------------------Number of Author Trend-----------------//
		String mean_acc_pro = map.get("mean_acc_pro");
		if (map.get("mean_acc_pro").length() > 8) {
			mean_acc_pro = map.get("mean_acc_pro").substring(0, 8);
		}

		String q1_acc_pro = map.get("q1_acc_pro");
		String q2_acc_pro = map.get("q2_acc_pro");
		String q3_acc_pro = map.get("q3_acc_pro");
		String Total_acc_pro = map.get("total_acc_pro");

		String mean_acc_kon = map.get("mean_acc_kontra");
		if (map.get("mean_acc_kontra").length() > 8) {
			mean_acc_kon = map.get("mean_acc_kontra").substring(0, 8);
		}

		String q1_acc_kon = map.get("q1_acc_kontra");
		String q2_acc_kon = map.get("q2_acc_kontra");
		String q3_acc_kon = map.get("q3_acc_kontra");
		String Total_acc_kon = map.get("total_acc_kontra");

		// ---------------Trend Comparison Kontra-------------//
		String pertumbuhan_tweet_kon = map.get("pertumbuhan_tweet_kontra") + "%";
		if (map.get("pertumbuhan_tweet_kontra").length() > 5) {
			pertumbuhan_tweet_kon = map.get("pertumbuhan_tweet_kontra").substring(0, 5) + "%";
		}
		String pertumbuhan_akun_kon = map.get("pertumbuhan_author_kontra") + "%";
		if (map.get("pertumbuhan_author_kontra").length() > 5) {
			pertumbuhan_akun_kon = map.get("pertumbuhan_author_kontra").substring(0, 5) + "%";
		}

		// ---------------Trend Comparison Pro-------------//
		String pertumbuhan_tweet_pro = map.get("pertumbuhan_tweet_pro") + "%";
		if (map.get("pertumbuhan_tweet_pro").length() > 6) {
			pertumbuhan_tweet_pro = map.get("pertumbuhan_tweet_pro").substring(0, 6) + "%";
		}
		String pertumbuhan_akun_pro = map.get("pertumbuhan_author_pro") + "%";
		if (map.get("pertumbuhan_author_pro").length() > 6) {
			pertumbuhan_akun_pro = map.get("pertumbuhan_author_pro").substring(0, 6) + "%";
		}

		// ----------Type of Post-----------//
		String tweet_terbanyak_pro = map.get("most_type_pro");
		String persen_tweet_terbanyak_pro = map.get("most_type_pro_per");
		String tweet_terbanyak_kon = map.get("most_type_kontra");
		String persen_tweet_terbanyak_kon = map.get("most_type_kontra_per");

		// ------------Topik Bahasan dan Kelompok-----------//
		String kata_terbanyak1_pro = map.get("most_word1_pro");
		String kata_terbanyak2_pro = map.get("most_word2_pro");
		String kata_terbanyak3_pro = map.get("most_word3_pro");
		String frekuensi_kata_terbanyak1_pro = map.get("most_word1_pro_count");
		String hashtag_terbanyak1_pro = "#" + map.get("most_tag1_pro");
		String hashtag_terbanyak2_pro = "#" + map.get("most_tag2_pro");
		String hashtag_terbanyak3_pro = "#" + map.get("most_tag3_pro");
		String frekuensi_hashtag_terbanyak1_pro = map.get("most_tag1_pro_count");
		String Mention_user1_pro = map.get("most_mention_pro");
		String Mention_user2_pro = "NULL";

		String kata_terbanyak1_kon = map.get("most_word1_kontra");
		String kata_terbanyak2_kon = map.get("most_word2_kontra");
		String kata_terbanyak3_kon = map.get("most_word3_kontra");
		String frekuensi_kata_terbanyak1_kon = map.get("most_word1_kontra_count");
		String hashtag_terbanyak1_kon = "#" + map.get("most_tag1_kontra");
		String hashtag_terbanyak2_kon = "#" + map.get("most_tag2_kontra");
		String hashtag_terbanyak3_kon = "#" + map.get("most_tag3_kontra");
		String frekuensi_hashtag_terbanyak1_kon = map.get("most_tag1_kontra_count");
		String Mention_user1_kon = map.get("most_mention_kontra");
		String Mention_user2_kon = "NULL";

		// ---------------TOP 10 ACCOUNT-------------//
		String account_by_follower1_kon = "@" + map.get("top1acc_byfollower_kontra");
		String account_by_reach1_kon = "@" + map.get("top1acc_byreach_kontra");

		String account_by_follower1_pro = "@" + map.get("top1acc_byfollower_pro");
		String account_by_reach1_pro = "@" + map.get("top1acc_byreach_pro");

		// // ===========Information========//
		// // gambar1
		// try {
		// saveImage(Url_image_akun1, "liker1TWv2.png");
		// } catch (Exception e) {
		// saveImage("https://pbs.twimg.com/profile_images/681560052008759296/defOyxnQ_200x200.png",
		// "liker1TWv2.png");
		// }
		//
		// // gambar2
		// try {
		// saveImage(Url_image_akun2, "liker2TWv2.png");
		// } catch (Exception e) {
		// saveImage("https://pbs.twimg.com/profile_images/681560052008759296/defOyxnQ_200x200.png",
		// "liker2TWv2.png");
		// }
		//
		// // gambar3
		// try {
		// saveImage(Url_image_akun3, "liker3TWv2.png");
		// } catch (Exception e) {
		// saveImage("https://pbs.twimg.com/profile_images/681560052008759296/defOyxnQ_200x200.png",
		// "liker3TWv2.png");
		// }

		// ======================SLIDE=====================//
		HSLFSlideShow ppt = new HSLFSlideShow();
		ppt.setPageSize(new java.awt.Dimension(width, height));

		// Slide 1
		HSLFSlide slide1 = ppt.createSlide();
		slideMaster(slide1);

		// Slide 2
		HSLFSlide slide2 = ppt.createSlide();
		slideInformation(ppt, slide2, time1, time2, akses, jumlah_data, Tweetakun1, Tweetakun2, Tweetakun3, Akun1,
				Akun2, Akun3);

		// Slide 3
		HSLFSlide slide3 = ppt.createSlide();
		slideTweetTrend(ppt, slide3, time1, time2, akses, mean_pro, q1_pro, q2_pro, q3_pro, Total_pro, mean_kon, q1_kon,
				q2_kon, q3_kon, Total_kon);

		// Slide 4
		HSLFSlide slide4 = ppt.createSlide();
		slideNumberofAuthorTrend(ppt, slide4, time1, time2, akses, mean_acc_pro, q1_acc_pro, q2_acc_pro, q3_acc_pro,
				Total_acc_pro, mean_acc_kon, q1_acc_kon, q2_acc_kon, q3_acc_kon, Total_acc_kon);

		// Slide 7
		HSLFSlide slide7 = ppt.createSlide();
		slideTypeofPostKontra(ppt, slide7, time1, time2, akses, tweet_terbanyak_kon, persen_tweet_terbanyak_kon);

		List<HSLFSlide> topretweet = new ArrayList<HSLFSlide>();
		for (int i = 0; i < 9; i++) {
			topretweet.add(ppt.createSlide());
			slideTopTweetRetweetsKontra(topretweet.get(i), time1, time2, akses, tokoh[i]);
			topretweet.get(i).addShape(tabeltweet.get(i));
			tabeltweet.get(i).moveTo(10, 100);
		}

		// Slide 12
		List<HSLFSlide> topbahasan = new ArrayList<HSLFSlide>();
		for (int i = 0; i < 9; i++) {
			topbahasan.add(ppt.createSlide());
			slideTopikBahasanKontra(ppt, topbahasan.get(i), time1, time2, akses, kata_terbanyak1_kon,
					kata_terbanyak2_kon, kata_terbanyak3_kon, frekuensi_kata_terbanyak1_kon, hashtag_terbanyak1_kon,
					hashtag_terbanyak2_kon, hashtag_terbanyak3_kon, frekuensi_hashtag_terbanyak1_kon, tokoh[i]);
			// ----- add picture-------------//
			HSLFPictureData pd = ppt.addPicture(
					new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_wordcloud.png"),
					PictureData.PictureType.PNG);
			HSLFPictureShape pictNew = new HSLFPictureShape(pd);
			pictNew.setAnchor(new java.awt.Rectangle(60, 100, 300, 300));
			topbahasan.get(i).addShape(pictNew);

			// ----- add picture-------------//
			HSLFPictureData pd1 = ppt.addPicture(
					new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_hashtagcloud.png"),
					PictureData.PictureType.PNG);
			HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
			pictNew1.setAnchor(new java.awt.Rectangle(540, 100, 300, 300));
			topbahasan.get(i).addShape(pictNew1);

		}

		// Slide 14.1
		HSLFSlide slide141 = ppt.createSlide();
		slideOverallNetwork(slide141, time1, time2, akses);

		// Slide 14
		List<HSLFSlide> topkelompok = new ArrayList<HSLFSlide>();
		for (int i = 0; i < 9; i++) {
			topkelompok.add(ppt.createSlide());
			slideKelompokKontra(ppt, topkelompok.get(i), time1, time2, akses, hashtag_terbanyak1_kon, Mention_user1_kon,
					tokoh[i]);
			// ----- add picture-------------//
			HSLFPictureData pd = ppt.addPicture(
					new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_TopUsedHashtag.png"),
					PictureData.PictureType.PNG);
			HSLFPictureShape pictNew = new HSLFPictureShape(pd);
			pictNew.setAnchor(new java.awt.Rectangle(40, 80, 420, 300));
			topkelompok.get(i).addShape(pictNew);

			// ----- add picture-------------//
			HSLFPictureData pd1 = ppt.addPicture(
					new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_TopMentionedUser.png"),
					PictureData.PictureType.PNG);
			HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
			pictNew1.setAnchor(new java.awt.Rectangle(500, 80, 420, 300));
			topkelompok.get(i).addShape(pictNew1);
		}

		// Slide 15
		List<HSLFSlide> top10Account = new ArrayList<HSLFSlide>();
		for (int i = 0; i < 9; i++) {
			top10Account.add(ppt.createSlide());
			slideTop10AccountKontra(top10Account.get(i), time1, time2, akses, account_by_follower1_kon,
					account_by_reach1_kon, tokoh[i]);
			top10Account.get(i).addShape(tabelbyfollower.get(i));
			tabelbyfollower.get(i).moveTo(50, 110);
			top10Account.get(i).addShape(tabelbyreach.get(i));
			tabelbyreach.get(i).moveTo(485, 110);
		}

		// Slide 17
		HSLFSlide slide17 = ppt.createSlide();
		slideTop10HashtagKontra(ppt, slide17, time1, time2, akses, akses);
		slide17.addShape(table_hashtag_growth_kon);
		table_hashtag_growth_kon.moveTo(50, 110);

		// Slide 19
		HSLFSlide slide19 = ppt.createSlide();
		slideNetworkofRetweet(slide19, time1, time2, akses);

		// Slide 20
		HSLFSlide slide20 = ppt.createSlide();
		slideTop10Active(ppt, slide20, time1, time2, akses);

		// Slide 21
		HSLFSlide slide21 = ppt.createSlide();
		slideNetworkofMention(slide21, time1, time2, akses);

		// Slide 22
		HSLFSlide slide22 = ppt.createSlide();
		slideKesimpulan(slide22, time1, time2, akses);

		FileOutputStream out = new FileOutputStream(filename);
		ppt.write(out);
		out.close();
		System.out.println("-------------------> Done <-------------------");
	}

	public static List<String> parseLine(String cvsLine) {
		return parseLine(cvsLine, DEFAULT_SEPARATOR, DEFAULT_QUOTE);
	}

	public static List<String> parseLine(String cvsLine, char separators) {
		return parseLine(cvsLine, separators, DEFAULT_QUOTE);
	}

	public static List<String> parseLine(String cvsLine, char separators, char customQuote) {

		List<String> result = new ArrayList<>();

		// if empty, return!
		if (cvsLine == null && cvsLine.isEmpty()) {
			return result;
		}

		if (customQuote == ' ') {
			customQuote = DEFAULT_QUOTE;
		}

		if (separators == ' ') {
			separators = DEFAULT_SEPARATOR;
		}

		StringBuffer curVal = new StringBuffer();
		boolean inQuotes = false;
		boolean startCollectChar = false;
		boolean doubleQuotesInColumn = false;

		char[] chars = cvsLine.toCharArray();

		for (char ch : chars) {

			if (inQuotes) {
				startCollectChar = true;
				if (ch == customQuote) {
					inQuotes = false;
					doubleQuotesInColumn = false;
				} else {

					// Fixed : allow "" in custom quote enclosed
					if (ch == '\"') {
						if (!doubleQuotesInColumn) {
							curVal.append(ch);
							doubleQuotesInColumn = true;
						}
					} else {
						curVal.append(ch);
					}

				}
			} else {
				if (ch == customQuote) {

					inQuotes = true;

					// Fixed : allow "" in empty quote enclosed
					if (chars[0] != '"' && customQuote == '\"') {
						curVal.append('"');
					}

					// double quotes in column will hit this!
					if (startCollectChar) {
						curVal.append('"');
					}

				} else if (ch == separators) {

					result.add(curVal.toString());

					curVal = new StringBuffer();
					startCollectChar = false;

				} else if (ch == '\r') {
					// ignore LF characters
					continue;
				} else if (ch == '\n') {
					// the end, break!
					break;
				} else {
					curVal.append(ch);
				}
			}

		}

		result.add(curVal.toString());

		return result;
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

	public HSLFTable tablev2(String csvFile, int a, int b) throws FileNotFoundException {

		int xx = 0;
		int xc = 0;
		if (b == 7 && a == 11) {
			xx = 1;
		} else if (b == 7 && a == 6) {
			xc = 1;
		}

		Scanner scanner = new Scanner(new File(csvFile));
		String[][] data = new String[a][b];
		int index = 0;
		String headers[] = new String[b];
		int k = 0;
		while (scanner.hasNext()) {
			List<String> line = parseLine(scanner.nextLine());
			if (index == 0) {
				for (int i = 0; i < b; i++) {
					headers[i] = line.get(i + xx);
				}
			} else {

				int ko = 0;
				k += 1;
				for (int c = 0 + xx; c < a + xc; c++) {
					try {
						data[k - 1][ko] = line.get(ko + xx);
						ko += 1;
					} catch (Exception e) {
					}
				}
			}
			index += 1;
		}
		scanner.close();

		HSLFTable table1 = new HSLFTable(a, b);
		for (int i = 0; i < a; i++) {
			for (int j = 0; j < b; j++) {
				HSLFTableCell cell = table1.getCell(i, j);
				HSLFTextRun rt = cell.getTextParagraphs().get(0).getTextRuns().get(0);
				rt.setFontFamily("Calibri");
				rt.setFontSize(12.);

				if (i == 0) {
					cell.getFill().setForegroundColor(Color.WHITE);
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
					cell.setFillColor(Color.white);
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
		border.setLineColor(new Color(255, 255, 255, 0));
		table1.setAllBorders(border);

		if (b == 5) {
			table1.setColumnWidth(0, 80);
			table1.setColumnWidth(1, 80);
			table1.setColumnWidth(2, 60);
			table1.setColumnWidth(4, 80);
			table1.setColumnWidth(3, 110);
		} else if (a == 11 && b == 7) {
			table1.setColumnWidth(0, 120);
		} else {
			table1.setColumnWidth(0, 80);
			table1.setColumnWidth(1, 60);
			table1.setColumnWidth(5, 100);
			table1.setColumnWidth(6, 400);

		}
		for (int i = 0; i < a; i++) {
			table1.setRowHeight(i, 20);

		}
		return table1;

	}

	public void Keterangan(HSLFSlide slide, int mode, String time1, String time2, String akses) {

		HSLFTextBox contentp = new HSLFTextBox();

		contentp.setText("Ket: Berdasarkan data sampel yang disediakan oleh twitter pada tanggal " + time1 + " sampai "
				+ time2 + ", diakses pada " + akses + ".");

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

	public void slideInformation(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses,
			String jumlah_data, String t1, String t2, String t3, String acc1, String acc2, String acc3)
			throws IOException {

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

	}

	public void slideTweetTrend(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses,
			String mean_pro, String q1p, String q2p, String q3p, String tp, String mean_kon, String q1k, String q2k,
			String q3k, String tk) throws IOException {

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
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\TweetFrequencyperDay.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(150, 90, 660, 310));
		slide.addShape(pictNew);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slideNumberofAuthorTrend(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses,
			String mean_pro, String q1p, String q2p, String q3p, String tp, String mean_kon, String q1k, String q2k,
			String q3k, String tk) throws IOException {

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
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\AuthorFrequencyperDay.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(150, 90, 660, 310));
		slide.addShape(pictNew);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slideTypeofPostKontra(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses,
			String tb, String pb) throws IOException {

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
		HSLFPictureData pd = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\AuthorFrequencyperDay.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(150, 90, 660, 310));
		slide.addShape(pictNew);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slideTopTweetRetweetsKontra(HSLFSlide slide, String time1, String time2, String akses, String nama) {

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

		content23.setText(nama);

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
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slideTopikBahasanKontra(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses,
			String kata1, String kata2, String kata3, String f1, String h1, String h2, String h3, String fh1,
			String nama) throws IOException {

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

		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText(nama);

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

		HSLFTextBox content23a = new HSLFTextBox();

		content23a.setText("Wordcloud\t\t\t\t\t\tHashtag");

		content23a.setAnchor(new java.awt.Rectangle(150, 380, 720, 50));

		HSLFTextParagraph contentP23a = content23a.getTextParagraphs().get(0);
		contentP23a.setAlignment(TextAlign.JUSTIFY);
		contentP23a.setLineSpacing(100.0);

		HSLFTextRun runContent23a = contentP23a.getTextRuns().get(0);
		runContent23a.setFontSize(18.);
		runContent23a.setFontFamily("Corbel");
		runContent23a.setBold(true);
		slide.addShape(content23a);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText(
				"Keywords yang paling banyak digunakan adalah \"" + kata1 + "\" dengan frekuensi penggunaan sebanyak "
						+ f1 + ". Disusul oleh keyword \"" + kata2 + "\" dan \"" + kata3 + "\".");

		content231.appendText(" Hashtag yang paling banyak digunakan adalah " + h1
				+ " dengan frekuensi penggunaan sebanyak " + fh1 + ". Disusul oleh hashtag " + h2 + " dan " + h3 + ".",
				false);

		content231.setAnchor(new java.awt.Rectangle(30, 410, 900, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(15.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);
	}

	public void slideOverallNetwork(HSLFSlide slide, String time1, String time2, String akses) {

		setBackground(slide, 1);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Overall Network");
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
		Keterangan(slide, 1, time1, time2, akses);
	}

	public void slideKelompokKontra(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses,
			String h1, String u1, String nama) throws IOException {

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

		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText(nama);

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

		content231.setText("Hashtag utama dalam kelompok kontra adalah " + h1
				+ " dan akun yang paling sering disebut oleh kelompok ini adalah " + u1 + ".");
		content231.setAnchor(new java.awt.Rectangle(50, 410, 860, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(14.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);
	}

	public void slideTop10AccountKontra(HSLFSlide slide, String time1, String time2, String akses, String af1,
			String ar1, String nama) {

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

		content23.setText(nama);

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

		content231.setText("Top Accounts Sorted By Follower\t\t\t\tTop Accounts Sorted By Reach");

		content231.setAnchor(new java.awt.Rectangle(50, 87, 810, 50));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(14.);
		runContent231.setFontFamily("Corbel");
		runContent231.setBold(true);
		slide.addShape(content231);

		// -------------------------Content------------------------//

		HSLFTextBox content231x = new HSLFTextBox();

		content231x.setText("Reach = jumlah tweet (count) X jumlah follower." + "\nDilihat dari jumlah follower, akun "
				+ af1
				+ " adalah akun yang memiliki paling banyak follower dalam kelompok akun-akun yang menggunakan #2019GantiPresiden dan tagar-tagar sejenis. Dilihat dari nilai reach, akun "
				+ ar1 + " merupakan akun yang memiliki nilai tertinggi.");
		content231x.setAnchor(new java.awt.Rectangle(50, 390, 860, 110));

		HSLFTextParagraph contentP231x = content231x.getTextParagraphs().get(0);
		contentP231x.setAlignment(TextAlign.JUSTIFY);
		contentP231x.setLineSpacing(100.0);

		HSLFTextRun runContent231x = contentP231x.getTextRuns().get(0);
		runContent231x.setFontSize(14.);
		runContent231x.setFontFamily("Corbel");

		slide.addShape(content231x);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);
	}

	public void slideTop10HashtagKontra(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses,
			String h) throws IOException {

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

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\posisi.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(770, 120, 170, 80));
		slide.addShape(pictNew);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText("Berikut adalah urutan hashtag berdasarkan pertumbuhan tertinggi pada periode " + time1
				+ " hingga " + time2
				+ " dari minggu sebelumnya.  Hashtag yang penggunaannya  meningkat paling tajam adalah " + h + ".");
		content231.setAnchor(new java.awt.Rectangle(50, 420, 860, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(14.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);
	}

	public void slideNetworkofRetweet(HSLFSlide slide, String time1, String time2, String akses) {

		setBackground(slide, 1);

		HSLFTextBox title = new HSLFTextBox();
		title.setText("Network of Retweet");
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
		Keterangan(slide, 1, time1, time2, akses);
	}

	public void slideTop10Active(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses)
			throws IOException {

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

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);
	}

	public void slideNetworkofMention(HSLFSlide slide, String time1, String time2, String akses) {

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
		Keterangan(slide, 1, time1, time2, akses);
	}

	public void slideKesimpulan(HSLFSlide slide, String time1, String time2, String akses) {

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
		SimpleDateFormat dateFormat = new SimpleDateFormat("dd MMMM yyyy");

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText("Berdasarkan data sampel yang disediakan oleh twitter pada tanggal " + time1 + " hingga "
				+ time2 + " diakses pada " + dateFormat.format(now) + ", dengan data sebagai berikut:");
		content231.setAnchor(new java.awt.Rectangle(50, 70, 860, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(18.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

	}

}
