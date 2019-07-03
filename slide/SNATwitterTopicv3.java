package com.ebdesk.report.slide;

import java.awt.Color;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.util.Date;
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

public class SNATwitterTopicv3 {

	private int version;
	private String filename;
	private int width;
	private int height;
	private String tema;
	private static final char DEFAULT_SEPARATOR = ',';
	private static final char DEFAULT_QUOTE = '"';

	public SNATwitterTopicv3(String filename, int version, int width, int height, String tema) {
		this.filename = filename;
		this.version = version;
		this.height = height;
		this.width = width;
		this.tema = tema;

	}

	public void execute() throws Exception {

		// =============TABEL============//
		HSLFTable table1pro = tablev2(
				"C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\pro_top10accountbyfollower.csv", 11, 5);
		HSLFTable table2pro = tablev2(
				"C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\pro_top10accountbyreach.csv", 11, 5);

		HSLFTable table1kontra = tablev2(
				"C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_top10accountbyfollower.csv", 11, 5);
		HSLFTable table2kontra = tablev2(
				"C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_top10accountbyreach.csv", 11, 5);

		HSLFTable table_hashtag_growth_pro = tablev2(
				"C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\pro_top10hastaggrowth.csv", 11, 7);

		HSLFTable table_hashtag_growth_kon = tablev2(
				"C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_top10hastaggrowth.csv", 11, 7);

		HSLFTable table_tweet_pro = tablev2(
				"C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\pro_top10tweet.csv", 6, 7);

		HSLFTable table_tweet_kon = tablev2(
				"C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_top10tweet.csv", 6, 7);

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

		// ===========Information========//
		// gambar1
		try {
			saveImage(Url_image_akun1, "liker1TWv2.png");
		} catch (Exception e) {
			saveImage("https://pbs.twimg.com/profile_images/681560052008759296/defOyxnQ_200x200.png", "liker1TWv2.png");
		}

		// gambar2
		try {
			saveImage(Url_image_akun2, "liker2TWv2.png");
		} catch (Exception e) {
			saveImage("https://pbs.twimg.com/profile_images/681560052008759296/defOyxnQ_200x200.png", "liker2TWv2.png");
		}

		// gambar3
		try {
			saveImage(Url_image_akun3, "liker3TWv2.png");
		} catch (Exception e) {
			saveImage("https://pbs.twimg.com/profile_images/681560052008759296/defOyxnQ_200x200.png", "liker3TWv2.png");
		}

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

		// Slide 5
		HSLFSlide slide5 = ppt.createSlide();
		slideTrendComparisonKontra(ppt, slide5, time1, time2, akses, tanggal_comparison_awal, tanggal_comparison_akhir,
				pertumbuhan_tweet_kon, pertumbuhan_akun_kon);

		// Slide 6
		HSLFSlide slide6 = ppt.createSlide();
		slideTrendComparisonPro(ppt, slide6, time1, time2, akses, tanggal_comparison_awal, tanggal_comparison_akhir,
				pertumbuhan_tweet_pro, pertumbuhan_akun_pro);

		// Slide 7
		HSLFSlide slide7 = ppt.createSlide();
		slideTypeofPostKontra(ppt, slide7, time1, time2, akses, tweet_terbanyak_kon, persen_tweet_terbanyak_kon);

		// Slide 8
		HSLFSlide slide8 = ppt.createSlide();
		slideTypeofPostPro(ppt, slide8, time1, time2, akses, tweet_terbanyak_pro, persen_tweet_terbanyak_pro);

		// Slide 91
		HSLFSlide slide91 = ppt.createSlide();
		slideTopTweetRetweetsKontra(slide91, time1, time2, akses);
		slide91.addShape(table_tweet_kon);
		table_tweet_kon.moveTo(10, 100);

		// Slide 101
		HSLFSlide slide101 = ppt.createSlide();
		slideTopTweetRetweetsPro(slide101, time1, time2, akses);
		slide101.addShape(table_tweet_pro);
		table_tweet_pro.moveTo(10, 100);

		// Slide 12
		HSLFSlide slide12 = ppt.createSlide();
		slideTopikBahasanKontra(ppt, slide12, time1, time2, akses, kata_terbanyak1_kon, kata_terbanyak2_kon,
				kata_terbanyak3_kon, frekuensi_kata_terbanyak1_kon, hashtag_terbanyak1_kon, hashtag_terbanyak2_kon,
				hashtag_terbanyak3_kon, frekuensi_hashtag_terbanyak1_kon);

		// Slide 11
		HSLFSlide slide11 = ppt.createSlide();
		slideTopikBahasanPro(ppt, slide11, time1, time2, akses, kata_terbanyak1_pro, kata_terbanyak2_pro,
				kata_terbanyak3_pro, frekuensi_kata_terbanyak1_pro, hashtag_terbanyak1_pro, hashtag_terbanyak2_pro,
				hashtag_terbanyak3_pro, frekuensi_hashtag_terbanyak1_pro);

		// Slide 14.1
		HSLFSlide slide141 = ppt.createSlide();
		slideOverallNetwork(slide141, time1, time2, akses);

		// Slide 14
		HSLFSlide slide14 = ppt.createSlide();
		slideKelompokKontra(ppt, slide14, time1, time2, akses, hashtag_terbanyak1_kon, Mention_user1_kon,
				Mention_user2_kon);
		// Slide 13
		HSLFSlide slide13 = ppt.createSlide();
		slideKelompokPro(ppt, slide13, time1, time2, akses, hashtag_terbanyak1_pro, Mention_user1_pro,
				Mention_user2_pro);

		// Slide 15
		HSLFSlide slide15 = ppt.createSlide();
		slideTop10AccountKontra(slide15, time1, time2, akses, account_by_follower1_kon, account_by_reach1_kon);
		slide15.addShape(table1kontra);
		table1kontra.moveTo(50, 110);
		slide15.addShape(table2kontra);
		table2kontra.moveTo(485, 110);

		// Slide 16
		HSLFSlide slide16 = ppt.createSlide();
		slideTop10AccountPro(slide16, time1, time2, akses, account_by_follower1_pro, account_by_reach1_pro);
		slide16.addShape(table1pro);
		table1pro.moveTo(50, 110);
		slide16.addShape(table2pro);
		table2pro.moveTo(485, 110);

		// Slide 17
		HSLFSlide slide17 = ppt.createSlide();
		slideTop10HashtagKontra(ppt, slide17, time1, time2, akses, akses);
		slide17.addShape(table_hashtag_growth_kon);
		table_hashtag_growth_kon.moveTo(50, 110);

		// Slide 18
		HSLFSlide slide18 = ppt.createSlide();
		slideTop10HashtagPro(ppt, slide18, time1, time2, akses, akses);
		slide18.addShape(table_hashtag_growth_pro);
		table_hashtag_growth_pro.moveTo(50, 110);

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
				+ jumlah_data + " data tweet.", false);
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

		content1.setText(t1);
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

		content2.setText(t2);
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

		content3.setText(t3);
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

		content12.setText(acc1);
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

		content22.setText(acc2);
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

		content23.setText(acc3);
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
		pictNew.setAnchor(new java.awt.Rectangle(20, 90, 630, 290));
		slide.addShape(pictNew);
		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Rata-rata tweet per hari\t:" + mean_pro);
		content23.appendText("\nKuartil1\t\t\t:" + q1p, false);
		content23.appendText("\nKuartil2\t\t\t:" + q2p, false);
		content23.appendText("\nKuartil3\t\t\t:" + q3p, false);
		content23.appendText("\nTotal Tweet\t\t:" + tp, false);
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

		content231.setText("Rata-rata tweet per hari\t:" + mean_kon);
		content231.appendText("\nKuartil1\t\t\t:" + q1k, false);
		content231.appendText("\nKuartil2\t\t\t:" + q2k, false);
		content231.appendText("\nKuartil3\t\t\t:" + q3k, false);
		content231.appendText("\nTotal Tweet\t\t:" + tk, false);
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
		pictNew.setAnchor(new java.awt.Rectangle(20, 90, 630, 290));
		slide.addShape(pictNew);

		// -------------------------Content------------------------//

		HSLFTextBox content23 = new HSLFTextBox();

		content23.setText("Rata-rata author per hari\t:" + mean_pro);
		content23.appendText("\nKuartil1\t\t\t:" + q1p, false);
		content23.appendText("\nKuartil2\t\t\t:" + q2p, false);
		content23.appendText("\nKuartil3\t\t\t:" + q3p, false);
		content23.appendText("\nTotal\t\t\t:" + tp, false);
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

		content231.setText("Rata-rata author per hari\t:" + mean_kon);
		content231.appendText("\nKuartil1\t\t\t:" + q1k, false);
		content231.appendText("\nKuartil2\t\t\t:" + q2k, false);
		content231.appendText("\nKuartil3\t\t\t:" + q3k, false);
		content231.appendText("\nTotal\t\t\t:" + tk, false);
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
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slideTrendComparisonKontra(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses,
			String tgl1, String tgl2, String ptk, String pak) throws IOException {

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
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_weeklytrendcomparison.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(50, 140, 400, 200));
		slide.addShape(pictNew);

		// ----- add picture-------------//
		HSLFPictureData pd1 = ppt.addPicture(
				new File(
						"C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_weeklytauthorcomparsion.png"),
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
				+ tgl1 + " hingga " + tgl2 + " mengalami pertumbuhan " + ptk + "dibandingkan minggu sebelumnya");
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

		content2312.setText("Jumlah akun yang menggunakan tagar-tagar bernada kontra kepada Jokowi pada rentang " + tgl1
				+ " hingga " + tgl2 + " mengalami pertumbuhan " + pak + " dibandingkan minggu sebelumnya");
		content2312.setAnchor(new java.awt.Rectangle(500, 400, 440, 110));

		HSLFTextParagraph contentP2312 = content2312.getTextParagraphs().get(0);
		contentP2312.setAlignment(TextAlign.JUSTIFY);
		contentP2312.setLineSpacing(100.0);

		HSLFTextRun runContent2312 = contentP2312.getTextRuns().get(0);
		runContent2312.setFontSize(14.);
		runContent2312.setFontFamily("Corbel");

		slide.addShape(content2312);

		// ========keterangan=========//
		Keterangan(slide, 2, tgl1, tgl2, akses);

	}

	public void slideTrendComparisonPro(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses,
			String tgl1, String tgl2, String ptp, String pap) throws IOException {

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
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\pro_weeklytrendcomparison.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(50, 140, 400, 200));
		slide.addShape(pictNew);

		// ----- add picture-------------//
		HSLFPictureData pd1 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\pro_weeklytauthorcomparsion.png"),
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

		content231.setText("Jumlah tweet yang menyertakan tagar bernada dukungan kepada Jokowi pada rentang " + tgl1
				+ " hingga " + tgl2 + " mengalami  pertumbuhan " + ptp + " dibandingkan minggu sebelumnya.");
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

		content2312.setText("Jumlah akun yang menggunakan tagar bernada dukungan kepada Jokowi pada rentang " + tgl1
				+ " hingga " + tgl2 + " mengalami pertumbuhan " + pap + " dibandingkan minggu sebelumnya.");
		content2312.setAnchor(new java.awt.Rectangle(500, 400, 440, 110));

		HSLFTextParagraph contentP2312 = content2312.getTextParagraphs().get(0);
		contentP2312.setAlignment(TextAlign.JUSTIFY);
		contentP2312.setLineSpacing(100.0);

		HSLFTextRun runContent2312 = contentP2312.getTextRuns().get(0);
		runContent2312.setFontSize(14.);
		runContent2312.setFontFamily("Corbel");

		slide.addShape(content2312);

		// ========keterangan=========//
		Keterangan(slide, 2, tgl1, tgl2, akses);

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
		HSLFPictureData pd1 = ppt.addPicture(new File(
				"C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_kontraPostTypesPercentages.png"),
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

		content231.setText("Tweet didominasi oleh aktivitas " + tb + " dengan persentase sebesar " + pb + ".");
		content231.setAnchor(new java.awt.Rectangle(655, 150, 300, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(16.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slideTypeofPostPro(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses,
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
		HSLFPictureData pd1 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\pro_proPostTypesPercentages.png"),
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

		content231.setText("Tweet didominasi oleh aktivitas " + tb + " dengan persentase sebesar " + pb + ".");
		content231.setAnchor(new java.awt.Rectangle(655, 150, 300, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(16.);
		runContent231.setFontFamily("Corbel");

		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);
	}

	public void slideTopTweetLikesKontra(HSLFSlide slide, String time1, String time2, String akses) throws IOException {

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
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slideTopTweetLikesPro(HSLFSlide slide, String time1, String time2, String akses) {

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
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slideTopTweetRetweetsKontra(HSLFSlide slide, String time1, String time2, String akses) {

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
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slideTopTweetRetweetsPro(HSLFSlide slide, String time1, String time2, String akses) {

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
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slideTopikBahasanKontra(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses,
			String kata1, String kata2, String kata3, String f1, String h1, String h2, String h3, String fh1)
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
		HSLFPictureData pd = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_wordcloud.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(60, 100, 300, 300));
		slide.addShape(pictNew);

		// ----- add picture-------------//
		HSLFPictureData pd1 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_hashtagcloud.png"),
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

		content231.setAnchor(new java.awt.Rectangle(320, 410, 600, 110));

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

	public void slideTopikBahasanPro(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses,
			String kata1, String kata2, String kata3, String f1, String h1, String h2, String h3, String fh1)
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
		HSLFPictureData pd = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\pro_wordcloud.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(60, 100, 300, 300));
		slide.addShape(pictNew);

		// ----- add picture-------------//
		HSLFPictureData pd1 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\pro_hashtagcloud.png"),
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

		content231.setAnchor(new java.awt.Rectangle(320, 410, 600, 110));

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

	public void slideOverallNetwork(HSLFSlide slide, String time1, String time2, String akses) {

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

	public void slideKelompokKontra(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses,
			String h1, String u1, String u2) throws IOException {

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
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_TopUsedHashtag.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(40, 80, 420, 300));
		slide.addShape(pictNew);

		// ----- add picture-------------//
		HSLFPictureData pd1 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\kontra_TopMentionedUser.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
		pictNew1.setAnchor(new java.awt.Rectangle(500, 80, 420, 300));
		slide.addShape(pictNew1);

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

	public void slideKelompokPro(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses,
			String h1, String u1, String u2) throws IOException {

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
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\pro_TopUsedHashtag.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(40, 80, 420, 300));
		slide.addShape(pictNew);

		// ----- add picture-------------//
		HSLFPictureData pd1 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\pro_TopMentionedUser.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
		pictNew1.setAnchor(new java.awt.Rectangle(500, 80, 420, 300));
		slide.addShape(pictNew1);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText("Hashtag utama dalam kelompok pro adalah " + h1
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
			String ar1) {

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

	public void slideTop10AccountPro(HSLFSlide slide, String time1, String time2, String akses, String af1,
			String ar1) {

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

		HSLFTextBox content231e = new HSLFTextBox();

		content231e.setText("Top Accounts Sorted By Follower\t\t\t\tTop Accounts Sorted By Reach");

		content231e.setAnchor(new java.awt.Rectangle(50, 87, 810, 50));

		HSLFTextParagraph contentP231e = content231e.getTextParagraphs().get(0);
		contentP231e.setAlignment(TextAlign.JUSTIFY);
		contentP231e.setLineSpacing(100.0);

		HSLFTextRun runContent231e = contentP231e.getTextRuns().get(0);
		runContent231e.setFontSize(14.);
		runContent231e.setBold(true);
		runContent231e.setFontFamily("Corbel");
		slide.addShape(content231e);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText("Reach = jumlah tweet (count) X jumlah follower." + "\nDilihat dari jumlah follower, akun "
				+ af1
				+ " adalah akun yang memiliki paling banyak follower dalam kelompok akun-akun yang menggunakan #2019tetapjokowi dan tagar-tagar sejenis. Dilihat dari nilai reach, akun "
				+ ar1 + " merupakan akun yang memiliki nilai tertinggi.");
		content231.setAnchor(new java.awt.Rectangle(50, 390, 860, 110));

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

	public void slideTop10HashtagPro(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses,
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
				+ "dari minggu sebelumnya.  Hashtag yang penggunaannya  meningkat paling tajam adalah " + h + ".");
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

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\pk\\pro_TopActiveUser.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(40, 100, 420, 340));
		slide.addShape(pictNew);

		// ----- add picture-------------//
		HSLFPictureData pd1 = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\kontra_TopActiveUser.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
		pictNew1.setAnchor(new java.awt.Rectangle(500, 100, 420, 340));
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
