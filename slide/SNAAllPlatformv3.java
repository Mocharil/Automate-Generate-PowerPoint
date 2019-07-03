package com.ebdesk.report.slide;

import java.awt.Color;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URL;
import java.util.Date;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Scanner;
import java.util.List;

import org.apache.commons.io.FileUtils;
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

public class SNAAllPlatformv3 {

	private int version;
	private String filename;
	private int width;
	private int height;
	private String tema;
	private String Platform;
	private static final char DEFAULT_SEPARATOR = ',';
	private static final char DEFAULT_QUOTE = '"';
	private String Post;
	private String Kirim;

	public SNAAllPlatformv3(int version, int width, int height, String tema, String Platform) {
		this.filename = Platform + "_" + tema + "v3.ppt";
		this.version = version;
		this.height = height;
		this.width = width;
		this.tema = tema;
		this.Platform = Platform;
		if (Platform.equals("Twitter")) {
			this.Post = "Tweet";
			this.Kirim = "Cuitan";
		} else {
			this.Post = "Post";
			this.Kirim = "Kiriman";
		}

	}

	public void execute() throws Exception {

		List<String> listnamapie = new ArrayList<String>();
		List<String> listnamatrend = new ArrayList<String>();
		List<String> listnamaauthor = new ArrayList<String>();
		List<String> listnamapost = new ArrayList<String>();
		List<String> listnamaword = new ArrayList<String>();
		List<String> listnamahcloud = new ArrayList<String>();
		List<String> listnamahashtag = new ArrayList<String>();
		List<String> listnamauser = new ArrayList<String>();
		List<String> listnamamention = new ArrayList<String>();
		List<String> listnamaurl = new ArrayList<String>();
		

		// ======PIE========//
		String[] png = new String[] { "png" };
		List<File> filepng = (List<File>) FileUtils.listFiles(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\" + tema + "_" + Platform), png, true);
		List<String> typepost = new ArrayList<String>();

		for (File file : filepng) {
			if (String.valueOf(file.getCanonicalPath()).toLowerCase().contains("post types percentages") == true) { // belum
				typepost.add(String.valueOf(file.getCanonicalPath()));
				Collections.sort(typepost, Collections.reverseOrder());
				try {
					listnamapie.add(String.valueOf(file.getName()).replace("_", "").replace(".png", "")
							.replace("post types percentages", "").replace("tokoh", "").replace(" ", ""));
				} catch (Exception e) {

				}
			}
		}

		// ======Trend Post========//

		List<String> trend = new ArrayList<String>();

		for (File file : filepng) {
			if (String.valueOf(file.getCanonicalPath()).toLowerCase().contains("postweeklycomparison") == true) { // belum
				trend.add(String.valueOf(file.getCanonicalPath()));
				Collections.sort(trend, Collections.reverseOrder());
				try {
					listnamatrend.add(String.valueOf(file.getName()).replace("_", "").replace(".png", "")
							.replace("postweeklycomparison", "").replace("tokoh", "").replace(" ", ""));
				} catch (Exception e) {

				}
			}

		}
		// ======Trend Author========//

		List<String> Author = new ArrayList<String>();

		for (File file : filepng) {
			if (String.valueOf(file.getCanonicalPath()).toLowerCase().contains("authorweeklycomparison") == true) { // belum
				Author.add(String.valueOf(file.getCanonicalPath()));
				Collections.sort(Author, Collections.reverseOrder());
				try {
					listnamaauthor.add(String.valueOf(file.getName()).replace("_", "").replace(".png", "")
							.replace("authorweeklycomparison", "").replace("tokoh", "").replace(" ", ""));
				} catch (Exception e) {

				}
			}

		}

		// ======BAR========//
		// url

		List<String> url = new ArrayList<String>();
		List<String> mention = new ArrayList<String>();
		List<String> user = new ArrayList<String>();
		List<String> hashtag = new ArrayList<String>();
		for (File file : filepng) {

			if (String.valueOf(file.getCanonicalPath()).contains("topsharedurl") == true) { // belum
				url.add(String.valueOf(file.getCanonicalPath()));
				try {
					listnamaurl.add(String.valueOf(file.getName()).replace("_", "").replace(".png", "")
							.replace("topsharedurl", "").replace("tokoh", "").replace(" ", ""));
				} catch (Exception e) {

				}
			}

			if (String.valueOf(file.getCanonicalPath()).contains("topusedhashtag") == true) {
				hashtag.add(String.valueOf(file.getCanonicalPath()));
				try {
					listnamahashtag.add(String.valueOf(file.getName()).replace("_", "").replace(".png", "")
							.replace("topusedhashtag", "").replace("tokoh", "").replace(" ", ""));
				} catch (Exception e) {

				}
			}
			if (String.valueOf(file.getCanonicalPath()).contains("topmentioneduser") == true) { // belum
				mention.add(String.valueOf(file.getCanonicalPath()));
				try {
					listnamamention.add(String.valueOf(file.getName()).replace("_", "").replace(".png", "")
							.replace("topmentioneduser", "").replace("tokoh", "").replace(" ", ""));
				} catch (Exception e) {

				}
			}
			if (String.valueOf(file.getCanonicalPath()).contains("topactiveuser") == true) {
				user.add(String.valueOf(file.getCanonicalPath()));
				try {
					listnamauser.add(String.valueOf(file.getName()).replace("_", "").replace(".png", "")
							.replace("topactiveuser", "").replace("tokoh", "").replace(" ", ""));
				} catch (Exception e) {

				}
			}
			Collections.sort(url, Collections.reverseOrder());
			Collections.sort(hashtag, Collections.reverseOrder());
			Collections.sort(mention, Collections.reverseOrder());
			Collections.sort(user, Collections.reverseOrder());
		}

		// =============TABEL============//

		String[] csv = new String[] { "csv" };
		List<File> filecsv = (List<File>) FileUtils.listFiles(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\" + tema + "_" + Platform), csv, true);
		List<HSLFTable> top10tweet = new ArrayList<HSLFTable>();
		List<HSLFTable> topGrowthtabel = new ArrayList<HSLFTable>();
		List<String> listtop10tweet = new ArrayList<String>();
		List<String> listGrowthtabel = new ArrayList<String>();
		for (File file : filecsv) {

			if (String.valueOf(file.getCanonicalPath()).contains("top post share") == true) { // belum

				listtop10tweet.add(String.valueOf(file.getCanonicalPath()));
				Collections.sort(listtop10tweet, Collections.reverseOrder());
				try {
					listnamapost.add(String.valueOf(file.getName()).replace("_", "").replace(".csv", "")
							.replace("top post share", "").replace("tokoh", "").replace(" ", ""));
				} catch (Exception e) {

				}
			}

			if (String.valueOf(file.getCanonicalPath()).contains("growth_overall") == true) { // belum

				listGrowthtabel.add(String.valueOf(file.getCanonicalPath()));
				Collections.sort(listGrowthtabel, Collections.reverseOrder());
			}

		}

		for (int i = 0; i < listtop10tweet.size(); i++) {
			top10tweet.add(tablev2(listtop10tweet.get(i), 6, 7));
		}
		for (int i = 0; i < listGrowthtabel.size(); i++) {
			topGrowthtabel.add(tablev2(listGrowthtabel.get(i), 11, 7));
		}

		// ======wordcloud========//
		List<String> wchtg = new ArrayList<String>();
		List<String> wc = new ArrayList<String>();
		for (File file : filepng) {

			if (String.valueOf(file.getCanonicalPath()).contains("wordcloud") == true) {

				wc.add(String.valueOf(file.getCanonicalPath()));

				try {
					listnamaword.add(String.valueOf(file.getName()).replace("_", "").replace(".png", "")
							.replace("wordcloud", "").replace("tokoh", "").replace(" ", ""));
				} catch (Exception e) {

				}

			}
			if (String.valueOf(file.getCanonicalPath()).contains("hashtagcloud") == true) {

				wchtg.add(String.valueOf(file.getCanonicalPath()));
				try {
					listnamahcloud.add(String.valueOf(file.getName()).replace("_", "").replace(".png", "")
							.replace("hashtagcloud", "").replace("tokoh", "").replace(" ", ""));
				} catch (Exception e) {

				}

			}
			Collections.sort(wc, Collections.reverseOrder());
			Collections.sort(wchtg, Collections.reverseOrder());
		}

		// ----------Type of Post-----------//
		List<String> tweet_terbanyak = new ArrayList<String>();
		List<String> persen_tweet_terbanyak = new ArrayList<String>();

		// ------------Topik Bahasan dan Kelompok-----------//

		List<String> kata_terbanyak1 = new ArrayList<String>();
		List<String> kata_terbanyak2 = new ArrayList<String>();
		List<String> kata_terbanyak3 = new ArrayList<String>();
		List<String> frekuensi_kata_terbanyak1 = new ArrayList<String>();

		List<String> hashtag_terbanyak1 = new ArrayList<String>();
		List<String> hashtag_terbanyak2 = new ArrayList<String>();
		List<String> hashtag_terbanyak3 = new ArrayList<String>();
		List<String> frekuensi_hashtag_terbanyak1 = new ArrayList<String>();

		List<String> User_terbanyak1 = new ArrayList<String>();
		List<String> User_terbanyak2 = new ArrayList<String>();
		List<String> User_terbanyak3 = new ArrayList<String>();

		List<String> pertumbuhan_hashtag1 = new ArrayList<String>();
		List<String> pertumbuhan_hashtag2 = new ArrayList<String>();
		List<String> pertumbuhan_hashtag3 = new ArrayList<String>();

		List<String> Mention_user1 = new ArrayList<String>();
		List<String> Mention_user2 = new ArrayList<String>();
		List<String> Mention_user3 = new ArrayList<String>();

		// ---------------TOP 10 ACCOUNT-------------//

		List<String> account_by_follower1 = new ArrayList<String>();
		List<String> account_by_follower2 = new ArrayList<String>();
		List<String> account_by_reach1 = new ArrayList<String>();
		List<String> account_by_reach2 = new ArrayList<String>();

		Collections.sort(listnamapie, Collections.reverseOrder());
		Collections.sort(listnamatrend, Collections.reverseOrder());
		Collections.sort(listnamaauthor, Collections.reverseOrder());
		Collections.sort(listnamapost, Collections.reverseOrder());
		Collections.sort(listnamaword, Collections.reverseOrder());
		Collections.sort(listnamahashtag, Collections.reverseOrder());
		Collections.sort(listnamauser, Collections.reverseOrder());
		Collections.sort(listnamamention, Collections.reverseOrder());
		Collections.sort(listnamaurl, Collections.reverseOrder());

		HashSet<String> setID = new HashSet<String>(listnamapie);
		setID.addAll(listnamatrend);
		setID.addAll(listnamaauthor);
		setID.addAll(listnamapost);
		setID.addAll(listnamaword);
		setID.addAll(listnamahashtag);
		setID.addAll(listnamauser);
		setID.addAll(listnamamention);
		setID.addAll(listnamaurl);
		setID.addAll(listnamahcloud);

		List<String> list = new ArrayList<String>(setID);
		Collections.sort(list, Collections.reverseOrder());

		// ====================CSV==================//
		String csvFile = "C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\" + tema + "_" + Platform
				+ "\\info.csv";
		Scanner scanner = new Scanner(new File(csvFile));

		Map<String, String> map = new HashMap<String, String>();
		List<String> line0 = new ArrayList<String>();
		int z = 0;
		while (scanner.hasNext()) {

			List<String> line = parseLine(scanner.nextLine());
			try {
				// if (z < 70) {
				map.put(line.get(0), line.get(1));
				line0.add(line.get(0));
				// map.put(line.get(0) + z, line.get(1));
				// line0.add(line.get(0) + z);
				z += 1;
			}

			catch (Exception e) {
			}
		}

		scanner.close();
		Collections.sort(line0, Collections.reverseOrder());

		for (int i = 0; i < map.size(); i++) {

			if (line0.get(i).contains("most_type_post_") && line0.get(i).contains("percentage") == false
					&& line0.get(i).contains("overall") == false) {
				tweet_terbanyak.add(map.get(line0.get(i)));

			}
			if (line0.get(i).contains("most_type_post_percetage")) {
				persen_tweet_terbanyak
						.add(String.valueOf(Float.valueOf(map.get(line0.get(i))) * 100).substring(0, 4) + "%");

			}

			if (line0.get(i).contains("most_tag_1")) {
				hashtag_terbanyak1.add("#" + map.get(line0.get(i)));

			}
			if (line0.get(i).contains("most_tag_freq_1")) {
				frekuensi_hashtag_terbanyak1.add(map.get(line0.get(i)));

			}
			if (line0.get(i).contains("most_tag_2")) {
				hashtag_terbanyak2.add("#" + map.get(line0.get(i)));

			}
			if (line0.get(i).contains("most_tag_3")) {
				hashtag_terbanyak3.add("#" + map.get(line0.get(i)));

			}
			if (line0.get(i).contains("most_mention_1")) {
				Mention_user1.add("@" + map.get(line0.get(i)));

			}
			if (line0.get(i).contains("most_mention_2")) {
				Mention_user2.add("@" + map.get(line0.get(i)));
			}
			if (line0.get(i).contains("most_mention_3")) {
				Mention_user3.add("@" + map.get(line0.get(i)));

			}
			if (line0.get(i).contains("most_word_1")) {
				kata_terbanyak1.add(map.get(line0.get(i)));

			}
			if (line0.get(i).contains("most_word_freq_1")) {
				frekuensi_kata_terbanyak1.add(map.get(line0.get(i)));

			}
			if (line0.get(i).contains("most_word_2")) {
				kata_terbanyak2.add(map.get(line0.get(i)));

			}
			if (line0.get(i).contains("most_word_3")) {
				kata_terbanyak3.add(map.get(line0.get(i)));

			}
			if (line0.get(i).contains("top1acc_byreach")) {
				account_by_reach1.add("@" + map.get(line0.get(i)));// belum

			}
			if (line0.get(i).contains("top1acc_byreach")) {
				account_by_reach2.add("@" + map.get(line0.get(i))); // belum

			}
			if (line0.get(i).contains("top1acc_byfollower")) {
				account_by_follower1.add("@" + map.get(line0.get(i))); // belum

			}
			if (line0.get(i).contains("top1acc_byfollower")) {
				account_by_follower2.add("@" + map.get(line0.get(i))); // belum

			}
			if (line0.get(i).contains("most_active_1")) {
				User_terbanyak1.add("@" + map.get(line0.get(i)));

			}
			if (line0.get(i).contains("most_active_2")) {
				User_terbanyak2.add("@" + map.get(line0.get(i)));

			}
			if (line0.get(i).contains("most_active_3")) {
				User_terbanyak3.add("@" + map.get(line0.get(i)));

			}

			if (line0.get(i).contains("most_hashtag_growth1_overall")) {
				pertumbuhan_hashtag1.add(map.get(line0.get(i)));
			}
			if (line0.get(i).contains("most_hashtag_growth2_overall")) {
				pertumbuhan_hashtag2.add(map.get(line0.get(i)));
			}
			if (line0.get(i).contains("most_hashtag_growth3_overall")) {
				pertumbuhan_hashtag3.add(map.get(line0.get(i)));
			}
		}

		// ================DATA================//

		String[] time = new String[5];
		time = map.get("min_date_overall").split("\\D");
		DateFormat df4 = new SimpleDateFormat("dd MM yyyy");
		DateFormat df = new SimpleDateFormat("dd MMMM yyyy");
		String waktu = new String();
		Date d1 = new Date();

		waktu = time[2] + " " + time[1] + " " + time[0];
		d1 = df4.parse(waktu);
		String time1 = df.format(d1); // tanggal awal

		time = map.get("max_date_overall").split("\\D");
		waktu = time[2] + " " + time[1] + " " + time[0];
		d1 = df4.parse(waktu);
		String time2 = df.format(d1);// tanggal akhir
		time = map.get("access_date_overall").split("\\D");
		waktu = time[2] + " " + time[1] + " " + time[0];
		d1 = df4.parse(waktu);
		String akses = df.format(d1); // tanggal akses

		// -------------INFO------------//
		String Akun1 = "@" + map.get("most_shared_account1_overall");
		String Akun2 = "@" + map.get("most_shared_account2_overall");
		String Akun3 = "@" + map.get("most_shared_account3_overall");
		String Tweetakun1 = null;
		try {
			Tweetakun1 = "\"" + map.get("most_shared_post1_overall").substring(0, 180) + "...\"";
		} catch (Exception e) {
			Tweetakun1 = "\"" + map.get("most_shared_post1_overall") + "\"";
		}
		String Tweetakun2 = null;
		try {
			Tweetakun2 = "\"" + map.get("most_shared_post2_overall").substring(0, 180) + "...\"";
		} catch (Exception e) {
			Tweetakun2 = "\"" + map.get("most_shared_post2_overall") + "\"";
		}
		String Tweetakun3 = null;
		try {
			Tweetakun3 = "\"" + map.get("most_shared_post3_overall").substring(0, 180) + "...\"";
		} catch (Exception e) {
			Tweetakun3 = "\"" + map.get("most_shared_post3_overall") + "\"";
		}
		String Url_image_akun1 = map.get("most_shared_img1_overall").replace("normal", "200x200");
		String Url_image_akun2 = map.get("most_shared_img2_overall").replace("normal", "200x200");
		String Url_image_akun3 = map.get("most_shared_img2_overall").replace("normal", "200x200");
		String jumlah_data = map.get("tweet_count_overall");
		String jumlah_akun = map.get("tweet_count"); // belum
		String jumlah_hashtag = map.get("tweet_count"); // belum

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

		// ----- add picture-------------//
		HSLFPictureData pdl = ppt.addPicture(
				new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\" + Platform + "_logo.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNewl = new HSLFPictureShape(pdl);
		pictNewl.setAnchor(new java.awt.Rectangle(25, 503, 30, 30));

		// Slide 1
		HSLFSlide slide1 = ppt.createSlide();
		slideMaster(ppt, slide1, time1, time2);
		slide1.addShape(pictNewl);

		// Slide 2
		HSLFSlide slide2 = ppt.createSlide();
		try {
			slideInformation(ppt, slide2, time1, time2, akses, jumlah_data, Tweetakun1, Tweetakun2, Tweetakun3, Akun1,
					Akun2, Akun3);
			slide2.addShape(pictNewl);
		} catch (Exception e) {
		}

		List<HSLFSlide> trendtw = new ArrayList<HSLFSlide>();
		for (int i = 0; i < trend.size(); i++) {
			// Slide TweetTrend
			trendtw.add(ppt.createSlide());
			slideTweetTrend(ppt, trendtw.get(i), time1, time2, akses);
			HSLFPictureData pd = ppt.addPicture(new File(trend.get(i)), PictureData.PictureType.PNG);
			HSLFPictureShape pictNew = new HSLFPictureShape(pd);
			pictNew.setAnchor(new java.awt.Rectangle(70, 100, 820, 350));
			trendtw.get(i).addShape(pictNew);
			trendtw.get(i).addShape(pictNewl);

			HSLFTextBox content23 = new HSLFTextBox();
			content23.setText(listnamatrend.get(i));
			content23.setAnchor(new java.awt.Rectangle(50, 20, 860, 75));

			HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
			contentP23.setAlignment(TextAlign.RIGHT);
			contentP23.setLineSpacing(100.0);

			HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
			runContent23.setFontSize(28.);
			runContent23.setItalic(true);
			runContent23.setFontFamily("Corbel");

			runContent23.setFontColor(new Color(255, 192, 0));
			trendtw.get(i).addShape(content23);

		}

		// Slide TweetTrendall
		if (new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\" + tema + "_" + Platform
				+ "\\postfrequencyperday.png").exists() == true) {
			HSLFSlide slide3 = ppt.createSlide();
			slideTweetTrendAll(ppt, slide3, time1, time2, akses);
			HSLFPictureData pd = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\" + tema
					+ "_" + Platform + "\\postfrequencyperday.png"), PictureData.PictureType.PNG);
			HSLFPictureShape pictNew = new HSLFPictureShape(pd);
			pictNew.setAnchor(new java.awt.Rectangle(70, 90, 820, 360));
			slide3.addShape(pictNew);
			slide3.addShape(pictNewl);
		}

		// Slide TweetTrend Boxplot
		if (new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\" + tema + "_" + Platform
				+ "\\post distribution per group.png").exists() == true) {
			HSLFSlide slide3 = ppt.createSlide();
			slideTweetTrenBox(ppt, slide3, time1, time2, akses);
			HSLFPictureData pd = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\" + tema
					+ "_" + Platform + "\\post distribution per group.png"), PictureData.PictureType.PNG);
			HSLFPictureShape pictNew = new HSLFPictureShape(pd);
			pictNew.setAnchor(new java.awt.Rectangle(70, 90, 820, 360));
			slide3.addShape(pictNew);
			slide3.addShape(pictNewl);
		}

		// Slide Number of Author
		List<HSLFSlide> author = new ArrayList<HSLFSlide>();
		for (int i = 0; i < Author.size(); i++) {

			author.add(ppt.createSlide());
			slideNumberofAuthorTrend(ppt, author.get(i), time1, time2, akses);
			HSLFPictureData pd = ppt.addPicture(new File(Author.get(i)), PictureData.PictureType.PNG);
			HSLFPictureShape pictNew = new HSLFPictureShape(pd);
			pictNew.setAnchor(new java.awt.Rectangle(70, 100, 820, 350));
			author.get(i).addShape(pictNew);
			author.get(i).addShape(pictNewl);

			HSLFTextBox content23 = new HSLFTextBox();
			content23.setText(listnamaauthor.get(i));
			content23.setAnchor(new java.awt.Rectangle(50, 20, 860, 75));

			HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
			contentP23.setAlignment(TextAlign.RIGHT);
			contentP23.setLineSpacing(100.0);

			HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
			runContent23.setFontSize(28.);
			runContent23.setItalic(true);
			runContent23.setFontFamily("Corbel");

			runContent23.setFontColor(new Color(255, 192, 0));
			author.get(i).addShape(content23);

		}

		// Slide Number of Authorall
		if (new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\" + tema + "_" + Platform
				+ "\\authorfrequencyperday.png").exists() == true) {
			HSLFSlide slide4 = ppt.createSlide();
			slideNumberofAuthorTrendAll(ppt, slide4, time1, time2, akses);
			HSLFPictureData pd = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\" + tema
					+ "_" + Platform + "\\authorfrequencyperday.png"), PictureData.PictureType.PNG);
			HSLFPictureShape pictNew = new HSLFPictureShape(pd);
			pictNew.setAnchor(new java.awt.Rectangle(70, 90, 820, 360));
			slide4.addShape(pictNew);
			slide4.addShape(pictNewl);
		}

		// Slide Number of Authora Boxplot
		if (new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\" + tema + "_" + Platform
				+ "\\author frequency distribution per group.png").exists() == true) {
			HSLFSlide slide4 = ppt.createSlide();
			slideNumberofAuthorTrendBox(ppt, slide4, time1, time2, akses);
			HSLFPictureData pd = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\" + tema
					+ "_" + Platform + "\\author frequency distribution per group.png"), PictureData.PictureType.PNG);
			HSLFPictureShape pictNew = new HSLFPictureShape(pd);
			pictNew.setAnchor(new java.awt.Rectangle(70, 90, 820, 360));
			slide4.addShape(pictNew);
			slide4.addShape(pictNewl);
		}

		// Slide Type of Post
		List<HSLFSlide> slide8 = new ArrayList<HSLFSlide>();
		for (int i = 0; i < typepost.size(); i++) {
			slide8.add(ppt.createSlide());
			slidePostType(ppt, slide8.get(i), time1, time2, akses, tweet_terbanyak.get(i),
					persen_tweet_terbanyak.get(i));
			// ----- add picture-------------//
			HSLFPictureData pd1 = ppt.addPicture(new File(typepost.get(i)), PictureData.PictureType.PNG);
			HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
			pictNew1.setAnchor(new java.awt.Rectangle(50, 100, 510, 370));
			slide8.get(i).addShape(pictNew1);
			slide8.get(i).addShape(pictNewl);

			HSLFTextBox content23 = new HSLFTextBox();
			content23.setText(listnamapie.get(i));
			content23.setAnchor(new java.awt.Rectangle(50, 20, 860, 75));

			HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
			contentP23.setAlignment(TextAlign.RIGHT);
			contentP23.setLineSpacing(100.0);

			HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
			runContent23.setFontSize(28.);
			runContent23.setItalic(true);
			runContent23.setFontFamily("Corbel");

			runContent23.setFontColor(Color.red);
			slide8.get(i).addShape(content23);
		}

		// slide Type of Post Bar
		if (new File(
				"C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\" + tema + "_" + Platform + "\\post type.png")
						.exists() == true) {
			HSLFSlide slide82 = ppt.createSlide();
			slidePostTypeAll(ppt, slide82, time1, time2, akses);
			// ----- add picture-------------//
			HSLFPictureData pd = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\" + tema
					+ "_" + Platform + "\\post type.png"), PictureData.PictureType.PNG);
			HSLFPictureShape pictNew = new HSLFPictureShape(pd);
			pictNew.setAnchor(new java.awt.Rectangle(100, 90, 760, 380));
			slide82.addShape(pictNew);
			slide82.addShape(pictNewl);

		}

		if (new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\" + tema + "_" + Platform
				+ "\\PostTypeComparison.png").exists() == true) {
			HSLFSlide slide82 = ppt.createSlide();
			slidePostTypeBox(ppt, slide82, time1, time2, akses);
			// ----- add picture-------------//
			HSLFPictureData pd = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\" + tema
					+ "_" + Platform + "\\PostTypeComparison.png"), PictureData.PictureType.PNG);
			HSLFPictureShape pictNew = new HSLFPictureShape(pd);
			pictNew.setAnchor(new java.awt.Rectangle(100, 90, 760, 380));
			slide82.addShape(pictNew);
			slide82.addShape(pictNewl);

		}

		// slide Top Tweet/Post
		List<HSLFSlide> top10 = new ArrayList<HSLFSlide>();
		for (int i = 0; i < top10tweet.size(); i++) {
			top10.add(ppt.createSlide());
			slideTopPost(top10.get(i), time1, time2, akses);
			top10.get(i).addShape(top10tweet.get(i));
			top10tweet.get(i).moveTo(10, 100);
			top10.get(i).addShape(pictNewl);

			HSLFTextBox content23 = new HSLFTextBox();
			content23.setText(listnamapost.get(i));
			content23.setAnchor(new java.awt.Rectangle(50, 20, 860, 75));

			HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
			contentP23.setAlignment(TextAlign.RIGHT);
			contentP23.setLineSpacing(100.0);

			HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
			runContent23.setFontSize(28.);
			runContent23.setItalic(true);
			runContent23.setFontFamily("Corbel");

			runContent23.setFontColor(new Color(255, 192, 0));
			top10.get(i).addShape(content23);

		}

		// ----------------------------------------------------------------------------------------///////////////////////
		// Slide Wordcloud and Hashtagcloud
		List<HSLFSlide> topbahasan = new ArrayList<HSLFSlide>();
		List<String> ww = new ArrayList<String>();
		HashSet<String> setword = new HashSet<String>(listnamaword);
		setword.addAll(listnamahcloud);
		ww.addAll(setword);
		Collections.sort(ww, Collections.reverseOrder());
		for (int k = 0; k < setword.size(); k++) {
			topbahasan.add(ppt.createSlide());
			// -------------------------KOTAK------------------------//
			HSLFTextBox content23ar = new HSLFTextBox();

			content23ar.setAnchor(new java.awt.Rectangle(375, 0, 585, 540));
			content23ar.setFillColor(new Color(204, 232, 234));
			if (Platform.equals("Instagram")) {
				content23ar.setFillColor(new Color(255, 200, 25));
			}
			topbahasan.get(k).addShape(content23ar);
		}

		for (int k = 0; k < list.size(); k++) {
			for (int i = 0; i < list.size(); i++) {
				try {
					if (wc.get(i).contains(ww.get(k)) || wchtg.get(i).contains(ww.get(k))) {

						setBackground(topbahasan.get(k), version);

						judul(topbahasan.get(k), "Topik Bahasan", "Kata-kata dan Hashtag yang paling banyak muncul");

						// ========keterangan=========//
						Keterangan(topbahasan.get(k), 2, time1, time2, akses);

						HSLFTextBox content23 = new HSLFTextBox();
						content23.setText(ww.get(k));

						content23.setAnchor(new java.awt.Rectangle(50, 20, 860, 75));

						HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
						contentP23.setAlignment(TextAlign.RIGHT);
						contentP23.setLineSpacing(100.0);

						HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
						runContent23.setFontSize(28.);
						runContent23.setItalic(true);
						runContent23.setFontFamily("Corbel");

						runContent23.setFontColor(Color.red);
						topbahasan.get(k).addShape(content23);
						topbahasan.get(k).addShape(pictNewl);

						try {
							if (wc.get(i).contains(ww.get(k))) {

								// -------------------------Content------------------------//

								HSLFTextBox content23a = new HSLFTextBox();

								content23a.setText("Wordcloud");

								content23a.setAnchor(new java.awt.Rectangle(380, 115, 200, 10));

								HSLFTextParagraph contentP23a = content23a.getTextParagraphs().get(0);
								contentP23a.setAlignment(TextAlign.JUSTIFY);
								contentP23a.setLineSpacing(120.0);

								HSLFTextRun runContent23a = contentP23a.getTextRuns().get(0);
								runContent23a.setFontSize(18.);
								runContent23a.setFontFamily("Corbel");
								runContent23a.setBold(true);
								topbahasan.get(k).addShape(content23a);

								// ----- add picture-------------//
								HSLFPictureData pd = ppt.addPicture(new File(wc.get(i)), PictureData.PictureType.PNG);
								HSLFPictureShape pictNew = new HSLFPictureShape(pd);
								pictNew.setAnchor(new java.awt.Rectangle(90, 90, 200, 200));
								topbahasan.get(k).addShape(pictNew);

								// -------------------------Content------------------------//
								int h = 0;
								HSLFTextBox content231 = new HSLFTextBox();
								for (int x = 0; x < map.size(); x++) {

									try {
										if (line0.get(x).contains("most_word_1_") && line0.get(x).contains(ww.get(k))) {

											content231.setText("Kata yang paling banyak digunakan pada topik " + tema
													+ " adalah ");
											content231.appendText(map.get(line0.get(x)), false);
											h = 5;

										}
									} catch (Exception e) {

									}

								}

								for (int z1 = 0; z1 < map.size(); z1++) {
									try {
										if (line0.get(z1).contains("most_word_2_")
												&& line0.get(z1).contains(ww.get(k))) {

											if (h == 5) {
												content231.appendText(". Disusul oleh kata ", false);
											} else {
												content231.setText(
														"Kata yang paling aktif pada topik " + tema + " adalah ");
												h = 6;
											}
											content231.appendText(map.get(line0.get(z1)), false);

										}
									} catch (Exception e) {
									}
								}

								for (int z2 = 0; z2 < map.size(); z2++) {
									try {
										if (line0.get(z2).contains("most_word_3_")
												&& line0.get(z2).contains(ww.get(k))) {

											if (h == 6) {
												content231.appendText(". Disusul oleh kata ", false);
											}

											else if (h == 5) {
												content231.appendText(" dan ", false);

											} else {
												content231.setText("Kata yang paling banyak digunakan pada topik "
														+ tema + " adalah ");
											}
											content231.appendText(map.get(line0.get(z2)) + ".", false);
										}
									} catch (Exception e) {
									}
								}

								content231.setAnchor(new java.awt.Rectangle(400, 135, 480, 110));
								HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
								contentP231.setAlignment(TextAlign.JUSTIFY);
								contentP231.setLineSpacing(100.0);

								HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
								runContent231.setFontSize(18.);
								runContent231.setFontFamily("Corbel");

								topbahasan.get(k).addShape(content231);

							}
						} catch (Exception e) {
						}
						try {
							if (wchtg.get(i).contains(ww.get(k))) {
								// -------------------------Content------------------------//

								HSLFTextBox content23a = new HSLFTextBox();

								content23a.setText("Hashtag");

								content23a.setAnchor(new java.awt.Rectangle(380, 285, 200, 10));

								HSLFTextParagraph contentP23a = content23a.getTextParagraphs().get(0);
								contentP23a.setAlignment(TextAlign.JUSTIFY);
								contentP23a.setLineSpacing(120.0);

								HSLFTextRun runContent23a = contentP23a.getTextRuns().get(0);
								runContent23a.setFontSize(18.);
								runContent23a.setFontFamily("Corbel");
								runContent23a.setBold(true);
								topbahasan.get(k).addShape(content23a);

								// ----- add picture-------------//
								HSLFPictureData pd1 = ppt.addPicture(new File(wchtg.get(i)),
										PictureData.PictureType.PNG);
								HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
								pictNew1.setAnchor(new java.awt.Rectangle(90, 290, 200, 200));
								topbahasan.get(k).addShape(pictNew1);

								// -------------------------Content------------------------//
								int h = 0;
								HSLFTextBox content231 = new HSLFTextBox();
								for (int x = 0; x < map.size(); x++) {

									try {
										if (line0.get(x).contains("most_tag_1_") && line0.get(x).contains(ww.get(k))) {

											content231.setText("Hashtag yang paling banyak digunakan pada topik " + tema
													+ " adalah ");
											content231.appendText(map.get(line0.get(x)), false);
											h = 5;

										}
									} catch (Exception e) {

									}

								}

								for (int z1 = 0; z1 < map.size(); z1++) {
									try {
										if (line0.get(z1).contains("most_tag_2_")
												&& line0.get(z1).contains(ww.get(k))) {
											if (h == 5) {
												content231.appendText(". Disusul oleh hashtag ", false);
											} else {
												content231.setText("Hashtag yang paling banyak digunakan pada topik "
														+ tema + " adalah ");
												h = 6;
											}
											content231.appendText(map.get(line0.get(z1)), false);

										}
									} catch (Exception e) {
									}
								}

								for (int z2 = 0; z2 < map.size(); z2++) {
									try {
										if (line0.get(z2).contains("most_tag_3_")
												&& line0.get(z2).contains(ww.get(k))) {
											if (h == 6) {

												content231.appendText(". Disusul oleh hashtag ", false);
											}

											else if (h == 5) {
												content231.appendText(" dan ", false);

											} else {
												content231.setText("Hashtag yang paling banyak digunakan pada topik "
														+ tema + " adalah ");
											}
											content231.appendText(map.get(line0.get(z2)) + ".", false);
										}
									} catch (Exception e) {
									}
								}

								content231.setAnchor(new java.awt.Rectangle(400, 320, 480, 110));
								HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
								contentP231.setAlignment(TextAlign.JUSTIFY);
								contentP231.setLineSpacing(100.0);

								HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
								runContent231.setFontSize(18.);
								runContent231.setFontFamily("Corbel");

								topbahasan.get(k).addShape(content231);

							}
						} catch (Exception e) {
						}
					}
				} catch (Exception e) {
				}

			}
		}

		// Slide Overall Network
		HSLFSlide slide21x = ppt.createSlide();
		slideOverallNetwork(slide21x, time1, time2, akses, jumlah_akun, jumlah_hashtag, jumlah_data);
		slide21x.addShape(pictNewl);

		// Slide Top 10 Hashtag and User or Mentioned User

		List<HSLFSlide> slidehashtag = new ArrayList<HSLFSlide>();
		List<String> jj = new ArrayList<String>();
		if (Platform.equals("Twitter")) {
			HashSet<String> sethashtagmention = new HashSet<String>(listnamamention);
			sethashtagmention.addAll(listnamahashtag);
			jj.addAll(sethashtagmention);
			Collections.sort(jj, Collections.reverseOrder());
			for (int k = 0; k < sethashtagmention.size(); k++) {
				slidehashtag.add(ppt.createSlide());
			}
		} else {
			HashSet<String> sethashtaguser = new HashSet<String>(listnamauser);
			sethashtaguser.addAll(listnamahashtag);
			jj.addAll(sethashtaguser);
			Collections.sort(jj, Collections.reverseOrder());
			for (int k = 0; k < sethashtaguser.size(); k++) {
				slidehashtag.add(ppt.createSlide());
			}
		}
	
		for (int k = 0; k < list.size(); k++) {
			for (int i = 0; i < list.size(); i++) {
				try {
					if (Platform.equals("Twitter")) {

						if (hashtag.get(i).contains(jj.get(k)) || mention.get(i).contains(jj.get(k))) {
							slidehashtagtwitter(ppt, slidehashtag.get(k), time1, time2, akses);

							HSLFTextBox content23 = new HSLFTextBox();
							content23.setText(jj.get(k));

							content23.setAnchor(new java.awt.Rectangle(50, 20, 860, 75));

							HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
							contentP23.setAlignment(TextAlign.RIGHT);
							contentP23.setLineSpacing(100.0);

							HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
							runContent23.setFontSize(28.);
							runContent23.setItalic(true);
							runContent23.setFontFamily("Corbel");

							runContent23.setFontColor(new Color(255, 192, 0));
							slidehashtag.get(k).addShape(content23);
							slidehashtag.get(k).addShape(pictNewl);
							try {
								if (hashtag.get(i).contains(jj.get(k))) {

									// ----- add picture-------------//
									HSLFPictureData pd = ppt.addPicture(new File(hashtag.get(i)),
											PictureData.PictureType.PNG);
									HSLFPictureShape pictNew = new HSLFPictureShape(pd);
									pictNew.setAnchor(new java.awt.Rectangle(40, 80, 420, 300));
									slidehashtag.get(k).addShape(pictNew);

									// -------------------------Content------------------------//
									int h = 0;
									HSLFTextBox content231 = new HSLFTextBox();
									for (int x = 0; x < map.size(); x++) {

										try {
											if (line0.get(x).contains("most_tag_1_")
													&& line0.get(x).contains(jj.get(k))) {

												content231.setText("Hashtag yang paling banyak digunakan pada topik "
														+ tema + " adalah ");
												content231.appendText("#" + map.get(line0.get(x)), false);
												h = 5;

											}
										} catch (Exception e) {

										}

									}

									for (int z1 = 0; z1 < map.size(); z1++) {
										try {
											if (line0.get(z1).contains("most_tag_2_")
													&& line0.get(z1).contains(jj.get(k))) {

												if (h == 5) {
													content231.appendText(". Disusul oleh hashtag ", false);
												} else {
													content231.setText("Hashtag yang paling aktif pada topik " + tema
															+ " adalah ");
													h = 6;
												}
												content231.appendText("#" + map.get(line0.get(z1)), false);

											}
										} catch (Exception e) {
										}
									}

									for (int z2 = 0; z2 < map.size(); z2++) {
										try {
											if (line0.get(z2).contains("most_tag_3_")
													&& line0.get(z2).contains(jj.get(k))) {

												if (h == 6) {
													content231.appendText(". Disusul oleh hashtag ", false);
												}

												else if (h == 5) {
													content231.appendText(" dan ", false);

												} else {
													content231
															.setText("Hashtag yang paling banyak digunakan pada topik "
																	+ tema + " adalah ");
												}
												content231.appendText("#" + map.get(line0.get(z2)) + ".", false);
											}
										} catch (Exception e) {
										}
									}

									content231.setAnchor(new java.awt.Rectangle(50, 410, 860, 110));
									HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
									contentP231.setAlignment(TextAlign.JUSTIFY);
									contentP231.setLineSpacing(100.0);

									HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
									runContent231.setFontSize(14.);
									runContent231.setFontFamily("Corbel");

									slidehashtag.get(k).addShape(content231);

								}
							} catch (Exception e) {
							}
							try {
								if (mention.get(i).contains(jj.get(k))) {
									// ----- add picture-------------//
									HSLFPictureData pd1 = ppt.addPicture(new File(mention.get(i)),
											PictureData.PictureType.PNG);
									HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
									pictNew1.setAnchor(new java.awt.Rectangle(500, 80, 420, 300));
									slidehashtag.get(k).addShape(pictNew1);

									// -------------------------Content------------------------//
									int h = 0;
									HSLFTextBox content231 = new HSLFTextBox();
									for (int x = 0; x < map.size(); x++) {

										try {
											if (line0.get(x).contains("most_mention_1_")
													&& line0.get(x).contains(jj.get(k))) {
												
												content231.setText("Akun yang paling banyak disebut pada topik " + tema
														+ " adalah ");
												content231.appendText(map.get(line0.get(x)), false);
												h = 5;

											}
										} catch (Exception e) {

										}

									}

									for (int z1 = 0; z1 < map.size(); z1++) {
										try {
											if (line0.get(z1).contains("most_mention_2_")
													&& line0.get(z1).contains(jj.get(k))) {
											
												if (h == 5) {
													content231.appendText(". Disusul oleh akun ", false);
												} else {
													content231.setText("Akun yang paling banyak disebut pada topik "
															+ tema + " adalah ");
													h = 6;
												}
												content231.appendText(map.get(line0.get(z1)), false);

											}
										} catch (Exception e) {
										}
									}

									for (int z2 = 0; z2 < map.size(); z2++) {
										try {
											if (line0.get(z2).contains("most_mention_3_")
													&& line0.get(z2).contains(jj.get(k))) {
									
												if (h == 6) {
													content231.appendText(". Disusul oleh akun ", false);
												}

												else if (h == 5) {
													content231.appendText(" dan ", false);

												} else {
													content231.setText("Akun yang paling banyak disebut pada topik "
															+ tema + " adalah ");
												}
												content231.appendText(map.get(line0.get(z2)) + ".", false);
											}
										} catch (Exception e) {
										}
									}

									content231.setAnchor(new java.awt.Rectangle(50, 440, 860, 110));
									HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
									contentP231.setAlignment(TextAlign.JUSTIFY);
									contentP231.setLineSpacing(100.0);

									HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
									runContent231.setFontSize(14.);
									runContent231.setFontFamily("Corbel");

									slidehashtag.get(k).addShape(content231);

								}
							} catch (Exception e) {
							}
						}

					}
					// ============== Hashtag User=============//
					else {
						if (hashtag.get(i).contains(jj.get(k)) || user.get(i).contains(jj.get(k))) {

							slidehashtaguser(ppt, slidehashtag.get(k), time1, time2, akses);
							HSLFTextBox content23 = new HSLFTextBox();
							content23.setText(jj.get(k));

							content23.setAnchor(new java.awt.Rectangle(50, 20, 860, 75));

							HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
							contentP23.setAlignment(TextAlign.RIGHT);
							contentP23.setLineSpacing(100.0);

							HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
							runContent23.setFontSize(28.);
							runContent23.setItalic(true);
							runContent23.setFontFamily("Corbel");

							runContent23.setFontColor(new Color(255, 192, 0));
							slidehashtag.get(k).addShape(content23);
							slidehashtag.get(k).addShape(pictNewl);

							try {
								if (hashtag.get(i).contains(jj.get(k))) {
									// ----- add picture-------------//
									HSLFPictureData pd = ppt.addPicture(new File(hashtag.get(i)),
											PictureData.PictureType.PNG);
									HSLFPictureShape pictNew = new HSLFPictureShape(pd);
									pictNew.setAnchor(new java.awt.Rectangle(40, 80, 420, 300));
									slidehashtag.get(k).addShape(pictNew);

									// -------------------------Content------------------------//
									int h = 0;
									HSLFTextBox content231 = new HSLFTextBox();
									for (int x = 0; x < map.size(); x++) {

										try {
											if (line0.get(x).contains("most_tag_1_")
													&& line0.get(x).contains(jj.get(k))) {

												content231.setText("Hashtag yang paling banyak digunakan pada topik "
														+ tema + " adalah ");
												content231.appendText("#" + map.get(line0.get(x)), false);
												h = 5;

											}
										} catch (Exception e) {

										}

									}

									for (int z1 = 0; z1 < map.size(); z1++) {
										try {
											if (line0.get(z1).contains("most_tag_2_")
													&& line0.get(z1).contains(jj.get(k))) {
												if (h == 5) {
													content231.appendText(". Disusul oleh hashtag ", false);
												} else {
													content231.setText("Hashtag yang paling aktif pada topik " + tema
															+ " adalah ");
													h = 6;
												}
												content231.appendText("#" + map.get(line0.get(z1)), false);

											}
										} catch (Exception e) {
										}
									}

									for (int z2 = 0; z2 < map.size(); z2++) {
										try {
											if (line0.get(z2).contains("most_tag_3_")
													&& line0.get(z2).contains(jj.get(k))) {
												if (h == 6) {
													content231.appendText(". Disusul oleh hashtag ", false);
												}

												else if (h == 5) {
													content231.appendText(" dan ", false);

												} else {
													content231
															.setText("Hashtag yang paling banyak digunakan pada topik "
																	+ tema + " adalah ");
												}
												content231.appendText("#" + map.get(line0.get(z2)) + ".", false);
											}
										} catch (Exception e) {
										}
									}

									content231.setAnchor(new java.awt.Rectangle(50, 410, 420, 110));
									HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
									contentP231.setAlignment(TextAlign.JUSTIFY);
									contentP231.setLineSpacing(100.0);

									HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
									runContent231.setFontSize(14.);
									runContent231.setFontFamily("Corbel");

									slidehashtag.get(k).addShape(content231);
								}
							} catch (Exception e) {
							}
							try {
								if (user.get(i).contains(jj.get(k))) {
									// ----- add picture-------------//
									HSLFPictureData pd1 = ppt.addPicture(new File(user.get(i)),
											PictureData.PictureType.PNG);
									HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
									pictNew1.setAnchor(new java.awt.Rectangle(500, 80, 420, 300));
									slidehashtag.get(k).addShape(pictNew1);

									// -------------------------Content------------------------//
									int h = 0;
									HSLFTextBox content231 = new HSLFTextBox();
									for (int x = 0; x < map.size(); x++) {

										try {
											if (line0.get(x).contains("most_active_1_")
													&& line0.get(x).contains(jj.get(k))) {

												content231.setText(
														"Akun yang paling aktif pada topik " + tema + " adalah ");
												content231.appendText("@" + map.get(line0.get(x)), false);
												h = 5;

											}
										} catch (Exception e) {

										}

									}

									for (int z1 = 0; z1 < map.size(); z1++) {
										try {
											if (line0.get(z1).contains("most_active_2_")
													&& line0.get(z1).contains(jj.get(k))) {
												if (h == 5) {
													content231.appendText(". Disusul oleh akun ", false);
												} else {
													content231.setText(
															"Akun yang paling aktif pada topik " + tema + " adalah ");
													h = 6;
												}
												content231.appendText("@" + map.get(line0.get(z1)), false);

											}
										} catch (Exception e) {
										}
									}

									for (int z2 = 0; z2 < map.size(); z2++) {
										try {
											if (line0.get(z2).contains("most_active_3_")
													&& line0.get(z2).contains(jj.get(k))) {
												if (h == 6) {
													content231.appendText(". Disusul oleh akun ", false);
												}

												else if (h == 5) {
													content231.appendText(" dan ", false);

												} else {
													content231.setText(
															"Akun yang paling aktif pada topik " + tema + " adalah ");
												}
												content231.appendText("@" + map.get(line0.get(z2)) + ".", false);
											}
										} catch (Exception e) {
										}
									}

									content231.setAnchor(new java.awt.Rectangle(500, 410, 420, 110));
									HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
									contentP231.setAlignment(TextAlign.JUSTIFY);
									contentP231.setLineSpacing(100.0);

									HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
									runContent231.setFontSize(14.);
									runContent231.setFontFamily("Corbel");

									slidehashtag.get(k).addShape(content231);

								}
							} catch (Exception e) {
							}
						}
					}
				} catch (Exception e) {
				}

			}
		}

		// slide TOP 10 User
		if (Platform.equals("Twitter")) {
			List<HSLFSlide> slideuser = new ArrayList<HSLFSlide>();

			for (int i = 0; i < user.size(); i++) {
				slideuser.add(ppt.createSlide());
				slideTopUser(ppt, slideuser.get(i), time1, time2, akses);
				// ----- add picture-------------//
				HSLFPictureData pd = ppt.addPicture(new File(user.get(i)), PictureData.PictureType.PNG);
				HSLFPictureShape pictNew = new HSLFPictureShape(pd);
				pictNew.setAnchor(new java.awt.Rectangle(40, 80, 490, 370));
				slideuser.get(i).addShape(pictNew);
				slideuser.get(i).addShape(pictNewl);

				HSLFTextBox content23 = new HSLFTextBox();
				content23.setText(listnamauser.get(i));
				content23.setAnchor(new java.awt.Rectangle(50, 20, 860, 75));

				HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
				contentP23.setAlignment(TextAlign.RIGHT);
				contentP23.setLineSpacing(100.0);

				HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
				runContent23.setFontSize(28.);
				runContent23.setItalic(true);
				runContent23.setFontFamily("Corbel");

				runContent23.setFontColor(Color.red);
				slideuser.get(i).addShape(content23);

				// -------------------------Content------------------------//
				int h = 0;
				HSLFTextBox content231 = new HSLFTextBox();
				for (int x = 0; x < map.size(); x++) {

					try {
						if (line0.get(x).contains("most_active_1_") && line0.get(x).contains(listnamauser.get(i))) {

							content231.setText("Akun yang paling aktif pada topik " + tema + " adalah ");
							content231.appendText("@"+map.get(line0.get(x)), false);
							h = 5;

						}
					} catch (Exception e) {

					}

				}

				for (int z1 = 0; z1 < map.size(); z1++) {
					try {
						if (line0.get(z1).contains("most_active_2_") && line0.get(z1).contains(listnamauser.get(i))) {
							if (h == 5) {
								content231.appendText(". Disusul oleh akun ", false);
							} else {
								content231.setText("Akun yang paling aktif pada topik " + tema + " adalah ");
								h = 6;
							}
							content231.appendText("@"+map.get(line0.get(z1)), false);

						}
					} catch (Exception e) {
					}
				}

				for (int z2 = 0; z2 < map.size(); z2++) {
					try {
						if (line0.get(z2).contains("most_active_3_") && line0.get(z2).contains(listnamauser.get(i))) {
							if (h == 6) {
								content231.appendText(". Disusul oleh akun ", false);
							}

							else if (h == 5) {
								content231.appendText(" dan ", false);

							} else {
								content231.setText("Link yang paling banyak di share pada topik " + tema + " adalah ");
							}
							content231.appendText("@"+map.get(line0.get(z2)) + ".", false);
						}
					} catch (Exception e) {
					}
				}

				content231.setAnchor(new java.awt.Rectangle(600, 150, 300, 110));
				HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
				contentP231.setAlignment(TextAlign.JUSTIFY);
				contentP231.setLineSpacing(100.0);

				HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
				runContent231.setFontSize(18.);
				runContent231.setFontFamily("Corbel");

				slideuser.get(i).addShape(content231);

			}

		}

		// slide Top 10 Url
		List<HSLFSlide> slideurl = new ArrayList<HSLFSlide>();
		for (int i = 0; i < url.size(); i++) {
			slideurl.add(ppt.createSlide());
			slideTopUrl(ppt, slideurl.get(i), time1, time2, akses);

			// ----- add picture-------------//
			HSLFPictureData pd = ppt.addPicture(new File(url.get(i)), PictureData.PictureType.PNG);
			HSLFPictureShape pictNew = new HSLFPictureShape(pd);
			pictNew.setAnchor(new java.awt.Rectangle(40, 80, 490, 370));
			slideurl.get(i).addShape(pictNew);
			slideurl.get(i).addShape(pictNewl);

			HSLFTextBox content23 = new HSLFTextBox();
			content23.setText(listnamaurl.get(i));
			content23.setAnchor(new java.awt.Rectangle(50, 20, 860, 75));

			HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
			contentP23.setAlignment(TextAlign.RIGHT);
			contentP23.setLineSpacing(100.0);

			HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
			runContent23.setFontSize(28.);
			runContent23.setItalic(true);
			runContent23.setFontFamily("Corbel");

			runContent23.setFontColor(Color.red);
			slideurl.get(i).addShape(content23);

			// -------------------------Content------------------------//
			int h = 0;
			HSLFTextBox content231 = new HSLFTextBox();
			for (int x = 0; x < map.size(); x++) {

				try {
					if (line0.get(x).contains("most_url_1_") && line0.get(x).contains(listnamaurl.get(i))) {

						content231.setText("Link yang paling banyak di share pada topik " + tema + " adalah ");
						content231.appendText(map.get(line0.get(x)), false);
						h = 5;

					}
				} catch (Exception e) {

				}

			}

			for (int z1 = 0; z1 < map.size(); z1++) {
				try {
					if (line0.get(z1).contains("most_url_2_") && line0.get(z1).contains(listnamaurl.get(i))) {
						if (h == 5) {
							content231.appendText(". Disusul oleh link ", false);
						} else {
							content231.setText("Link yang paling banyak di share pada topik " + tema + " adalah ");
							h = 6;
						}
						content231.appendText(map.get(line0.get(z1)), false);

					}
				} catch (Exception e) {
				}
			}

			for (int z2 = 0; z2 < map.size(); z2++) {
				try {
					if (line0.get(z2).contains("most_url_3_") && line0.get(z2).contains(listnamaurl.get(i))) {
						if (h == 6) {
							content231.appendText(". Disusul oleh link ", false);
						}

						else if (h == 5) {
							content231.appendText(" dan ", false);

						} else {
							content231.setText("Link yang paling banyak di share pada topik " + tema + " adalah ");
						}
						content231.appendText(map.get(line0.get(z2)) + ".", false);
					}
				} catch (Exception e) {
				}
			}

			content231.setAnchor(new java.awt.Rectangle(600, 150, 300, 110));
			HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
			contentP231.setAlignment(TextAlign.JUSTIFY);
			contentP231.setLineSpacing(100.0);

			HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
			runContent231.setFontSize(18.);
			runContent231.setFontFamily("Corbel");

			slideurl.get(i).addShape(content231);

		}

		// slide Top 10 account by follower and reach
		if (Platform.equals("Twitter")) {

			List<File> files10a = (List<File>) FileUtils.listFiles(
					new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\" + tema + "_" + Platform), csv,
					true);
			List<HSLFTable> top10afollower = new ArrayList<HSLFTable>();
			List<HSLFTable> top10areach = new ArrayList<HSLFTable>();
			List<String> listtop10afollower = new ArrayList<String>();
			List<String> listtop10areach = new ArrayList<String>();
			HashSet<String> settop = new HashSet<String>();
			for (File file : files10a) {
				if (String.valueOf(file.getCanonicalPath()).contains("follower") == true) {
					listtop10afollower.add(file.getCanonicalPath());
				
					try {
						settop.add(String.valueOf(file.getName()).replace("_", "").replace(".csv", "")
								.replace("follower", "").replace("tokoh", "").replace(" ", ""));
					} catch (Exception e) {

					}
				}
				if (String.valueOf(file.getCanonicalPath()).contains("reach") == true) {
					listtop10areach.add(file.getCanonicalPath());
					Collections.sort(listtop10areach, Collections.reverseOrder());
					try {
						settop.add(String.valueOf(file.getName()).replace("_", "").replace(".csv", "")
								.replace("reach", "").replace("tokoh", "").replace(" ", ""));
						
					} catch (Exception e) {

					}
				}
			}
			List<String> listnamatop = new ArrayList<String>(settop);
			Collections.sort(listnamatop, Collections.reverseOrder());
			for (int i = 0; i < listtop10afollower.size(); i++) {
				top10afollower.add(tablev2(listtop10afollower.get(i), 11, 5));
			}

			for (int i = 0; i < listtop10areach.size(); i++) {
				top10areach.add(tablev2(listtop10areach.get(i), 11, 5));
			}
			
			List<HSLFSlide> top10a = new ArrayList<HSLFSlide>();
			for (int i = 0; i < Math.max(top10afollower.size(), top10areach.size()); i++) {
				top10a.add(ppt.createSlide());
				slideTop10Accounts(top10a.get(i), time1, time2, akses);
				top10a.get(i).addShape(top10afollower.get(i));
				top10afollower.get(i).moveTo(50, 110);
				top10a.get(i).addShape(top10areach.get(i));
				top10areach.get(i).moveTo(485, 110);
				top10a.get(i).addShape(pictNewl);

				HSLFTextBox content23 = new HSLFTextBox();
				content23.setText(listnamatop.get(i));

				content23.setAnchor(new java.awt.Rectangle(50, 20, 860, 75));

				HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
				contentP23.setAlignment(TextAlign.RIGHT);
				contentP23.setLineSpacing(100.0);

				HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
				runContent23.setFontSize(28.);
				runContent23.setItalic(true);
				runContent23.setFontFamily("Corbel");

				runContent23.setFontColor(new Color(255, 192, 0));
				top10a.get(i).addShape(content23);

				// -------------------------Content------------------------//
				int h = 0;
				int a = 0;
				HSLFTextBox content231 = new HSLFTextBox();
				for (int x = 0; x < map.size(); x++) {

					try {
						if (line0.get(x).contains("top1acc_byfollower") && line0.get(x).contains(listnamatop.get(i))) {

							content231.setText("Reach = jumlah tweet (count) X jumlah follower."
									+ "\nDilihat dari jumlah follower, akun ");
							content231.appendText(map.get(line0.get(x)), false);
							content231.appendText(
									" adalah akun yang memiliki paling banyak follower pada topik " + tema + "ini.",
									false);
							h = 5;

						}
					} catch (Exception e) {

					}

				}

				for (int z1 = 0; z1 < map.size(); z1++) {
					try {
						if (line0.get(z1).contains("top2acc_byfollower")
								&& line0.get(z1).contains(listnamatop.get(i))) {
							if (h == 5) {
								content231.appendText(". Disusul oleh akun ", false);
							} else {
								content231.setText("Reach = jumlah tweet (count) X jumlah follower."
										+ "\nDilihat dari jumlah follower, akun ");
								h = 6;
							}
							content231.appendText(map.get(line0.get(z1)), false);
							if (h == 6) {
								content231.appendText(
										" adalah akun yang memiliki paling banyak follower pada topik " + tema + "ini.",
										false);
							}

						}
					} catch (Exception e) {
					}
				}

				for (int z2 = 0; z2 < map.size(); z2++) {
					try {
						if (line0.get(z2).contains("top1acc_byreach") && line0.get(z2).contains(listnamatop.get(i))) {

							content231.appendText(". Dilihat dari nilai reach, akun ", false);
							content231.appendText(map.get(line0.get(z2)) + ".", false);
							content231.appendText(" merupakan akun yang memiliki nilai tertinggi.", false);
							a = 6;
						}
					} catch (Exception e) {
					}
				}

				for (int z2 = 0; z2 < map.size(); z2++) {
					try {
						if (line0.get(z2).contains("top2acc_byreach") && line0.get(z2).contains(listnamatop.get(i))) {
							if (a != 6) {
								content231.appendText(". Dilihat dari nilai reach, akun ", false);
							}

							else {
								content231.appendText(". Disusul oleh akun ", false);
							}
							content231.appendText(map.get(line0.get(z2)) + ".", false);
							if (a != 6) {
								content231.appendText(" merupakan akun yang memiliki nilai tertinggi.", false);
							}
						}
					} catch (Exception e) {
					}
				}

				content231.setAnchor(new java.awt.Rectangle(50, 390, 860, 110));
				HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
				contentP231.setAlignment(TextAlign.JUSTIFY);
				contentP231.setLineSpacing(100.0);

				HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
				runContent231.setFontSize(18.);
				runContent231.setFontFamily("Corbel");

				top10a.get(i).addShape(content231);

			}
		}

		// Slide Top 10 Hashtag Growth
		List<HSLFSlide> topGrowth = new ArrayList<HSLFSlide>();
		for (int i = 0; i < topGrowthtabel.size(); i++) {
			topGrowth.add(ppt.createSlide());
			slideTop10HashtagGrowth(ppt, topGrowth.get(i), time1, time2, akses);
			topGrowth.get(i).addShape(topGrowthtabel.get(i));
			topGrowthtabel.get(i).moveTo(30, 100);
			topGrowth.get(i).addShape(pictNewl);

			// -------------------------Content------------------------//

			HSLFTextBox content231 = new HSLFTextBox();
			int h = 0;
			content231.setText("Berikut adalah urutan hashtag berdasarkan pertumbuhan tertinggi pada periode " + time1
					+ " hingga " + time2
					+ "dari minggu sebelumnya.  Hashtag yang penggunaannya  meningkat paling tajam adalah ");
			try {
				content231.appendText(pertumbuhan_hashtag1.get(i), false);
				h = 5;
			} catch (Exception e) {
				break;
			}
			try {
				if (h == 5 && pertumbuhan_hashtag2.get(i).length() > 1) {
					content231.appendText(". Disusul oleh hashtag ", false);
					content231.appendText(pertumbuhan_hashtag2.get(i), false);
					h = 5;
				}
			} catch (Exception e) {
				break;
			}
			try {
				if (h == 5 && pertumbuhan_hashtag3.get(i).length() > 1) {
					content231.appendText(" dan ", false);
					content231.appendText(pertumbuhan_hashtag3.get(i) + ".", false);
				}
			} catch (Exception e) {
			}

			content231.setAnchor(new java.awt.Rectangle(50, 420, 860, 110));

			HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
			contentP231.setAlignment(TextAlign.JUSTIFY);
			contentP231.setLineSpacing(100.0);

			HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
			runContent231.setFontSize(14.);
			runContent231.setFontFamily("Corbel");

			topGrowth.get(i).addShape(content231);

		}

		// Slide Network of Retweet or Shares
		HSLFSlide slide19 = ppt.createSlide();
		slideNetworkofRetweet(slide19, time1, time2, akses);
		slide19.addShape(pictNewl);

		// Slide Network of Mention or Post
		HSLFSlide slide21 = ppt.createSlide();
		slideNetworkofMention(slide21, time1, time2, akses);
		slide21.addShape(pictNewl);

		// Slide Kesimpulan
		HSLFSlide slide22 = ppt.createSlide();
		slideKesimpulan(slide22, time1, time2, akses);
		slide22.addShape(pictNewl);

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

	public void saveImage(String imageUrl, String destinationFile) throws IOException {
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

		BufferedReader br = null;
		String line = "";
		String cvsSplitBy = ",";

		int xx = 0;
		int xc = 0;
		if (b == 7 && a == 11) {
			xx = 0;
		} else if (b == 7 && a == 6) {
			xc = 1;
		}

		String[][] data = new String[a][b];
		int index = 0;
		String headers[] = new String[b];
		int k = 0;

		try {

			br = new BufferedReader(new FileReader(csvFile));
			while ((line = br.readLine()) != null) {

				// use comma as separator
				String[] isi = line.split(cvsSplitBy);

				// ==========//
				if (index == 0) {
					for (int i = 0; i < b; i++) {
						headers[i] = isi[i + xx];
					}
				} else {

					int ko = 0;
					k += 1;
					for (int c = 0 + xx; c < a + xc; c++) {
						try {
							if (isi[ko + xx].length() >= 140) {
								data[k - 1][ko] = isi[ko + xx].substring(0, 140) + "...";
							} else {
								data[k - 1][ko] = isi[ko + xx];
							}
							ko += 1;
						} catch (Exception e) {

						}
					}
				}
				index += 1;

				// ==========
			}

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (br != null) {
				try {
					br.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}

		HSLFTable table1 = new HSLFTable(a, b);
		for (int i = 0; i < a; i++) {
			for (int j = 0; j < b; j++) {
				HSLFTableCell cell = table1.getCell(i, j);
				HSLFTextRun rt = cell.getTextParagraphs().get(0).getTextRuns().get(0);
				rt.setFontFamily("Calibri");
				rt.setFontSize(12.);

				if (i == 0) {
					cell.getFill().setForegroundColor(new Color(255, 255, 255, 0));
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
					cell.setFillColor(new Color(255, 255, 255, 0));
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
			table1.setColumnWidth(1, 80);
			table1.setColumnWidth(2, 120);
			table1.setColumnWidth(5, 80);
			table1.setColumnWidth(6, 385);

		}
		for (int i = 0; i < a; i++) {
			table1.setRowHeight(i, 20);

		}
		return table1;

	}
	public void Platform(HSLFSlide slide) {

		HSLFTextBox contentp = new HSLFTextBox();

		contentp.setText("Platform: " + Platform);

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
		runTitlep.setFontColor(new Color(61, 140, 255));
		if (Platform.equals("Instagram")) {
			runTitlep.setFontColor(new Color(255, 192, 0));
		}

		slide.addShape(contentp);
	}

	public void Garis(HSLFSlide slide, int a, int b, int c, int d) {
		HSLFTextBox content = new HSLFTextBox();
		content.setAnchor(new java.awt.Rectangle(a, b, c, d));
		content.setLineColor(new Color(202, 202, 202));
		content.setLineWidth(1);
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
		runTitle.setFontColor(new Color(61, 140, 255));
		if (Platform.equals("Instagram")) {
			runTitle.setFontColor(new Color(255, 192, 0));
		}

		slide.addShape(title);

		// --------------------Investigation--------------------
		HSLFTextBox issue = new HSLFTextBox();
		issue.setText(tema);
		issue.setAnchor(new java.awt.Rectangle(25, 50, 930, 50));
		HSLFTextParagraph issueP = issue.getTextParagraphs().get(0);
		HSLFTextRun runIssue = issueP.getTextRuns().get(0);
		runIssue.setFontSize(17.);
		runIssue.setFontFamily("Calibri");
		runIssue.setBold(false);
		if (atas.contains("Network")) {
			runIssue.setFontColor(Color.white);
		} else {
			runIssue.setFontColor(Color.black);
		}
		slide.addShape(issue);

		// =========garis===========//
		Garis(slide, 30, 78, 100, 1);

		// =======content Platform========//
		Platform(slide);

		HSLFTextBox contentx = new HSLFTextBox();
		contentx.setAnchor(new java.awt.Rectangle(25, 30, 0, 45));
		contentx.setLineColor(new Color(0, 136, 184));
		if (Platform.equals("Instagram")) {
			contentx.setLineColor(new Color(255, 192, 0));
		}
		contentx.setRotation(180);
		contentx.setLineWidth(4);
		slide.addShape(contentx);

	}

	public void Keterangan(HSLFSlide slide, int mode, String time1, String time2, String akses) {

		HSLFTextBox contentp = new HSLFTextBox();

		contentp.setText("Ket: Berdasarkan data sampel yang disediakan oleh " + Platform + " pada tanggal " + time1
				+ " sampai " + time2 + ", diakses pada " + akses + ".");

		contentp.setAnchor(new java.awt.Rectangle(605, 495, 350, 40));
		contentp.setLineColor(new Color(204, 232, 234));
		if (Platform.equals("Instagram")) {
			contentp.setLineColor(new Color(255, 200, 25));
		}
		contentp.setLineWidth(2);

		HSLFTextParagraph titlePp = contentp.getTextParagraphs().get(0);
		titlePp.setAlignment(TextAlign.JUSTIFY);
		titlePp.setSpaceAfter(0.);
		titlePp.setSpaceBefore(0.);
		titlePp.setLineSpacing(110.0);

		HSLFTextRun runTitlep = titlePp.getTextRuns().get(0);
		runTitlep.setFontSize(11.);
		runTitlep.setFontFamily("Corbel");
		if (mode == 1) {
			runTitlep.setFontColor(Color.white);
		}

		slide.addShape(contentp);
	}

	public void slideMaster(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2) throws IOException {
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
		runTitle.setFontFamily("Bell MT");
		runTitle.setBold(true);
		runTitle.setUnderlined(false);
		runTitle.setFontColor(new Color(61, 140, 255));
		if (Platform.equals("Instagram")) {
			runTitle.setFontColor(new Color(255, 192, 0));
		}

		slide.addShape(title);

		// --------------------Investigation--------------------
		HSLFTextBox issue = new HSLFTextBox();
		issue.setText("Analis Topic " + tema + " : " + time1 + " - " + time2);
		issue.setAnchor(new java.awt.Rectangle(10, 270, 950, 50));
		HSLFTextParagraph issueP = issue.getTextParagraphs().get(0);
		issueP.setAlignment(TextAlign.CENTER);
		issueP.setSpaceAfter(0.);
		issueP.setSpaceBefore(0.);

		HSLFTextRun runIssue = issueP.getTextRuns().get(0);
		runIssue.setFontSize(18.);
		runIssue.setFontFamily("Calibri");
		runIssue.setBold(false);
		runIssue.setFontColor(new Color(166, 166, 166));
		slide.addShape(issue);

		// =========garis===========//
		Garis(slide, 430, 308, 100, 1);

		// =======content Platform========//
		Platform(slide);

		// =====================garis====================//
		HSLFTextBox contentx = new HSLFTextBox();
		contentx.setAnchor(new java.awt.Rectangle(220, 230, 45, 0));
		contentx.setLineColor(new Color(0, 176, 240));
		if (Platform.equals("Instagram")) {
			contentx.setLineColor(new Color(255, 192, 0));
		}
		contentx.setRotation(180);
		contentx.setLineWidth(4);
		slide.addShape(contentx);

		HSLFTextBox contentx1 = new HSLFTextBox();
		contentx1.setAnchor(new java.awt.Rectangle(220, 230, 0, 45));
		contentx1.setLineColor(new Color(0, 176, 240));
		if (Platform.equals("Instagram")) {
			contentx1.setLineColor(new Color(255, 192, 0));
		}
		contentx1.setLineWidth(4);
		slide.addShape(contentx1);

		HSLFTextBox contentx2 = new HSLFTextBox();
		contentx2.setAnchor(new java.awt.Rectangle(710, 315, 45, 0));
		contentx2.setLineColor(new Color(0, 176, 240));
		if (Platform.equals("Instagram")) {
			contentx2.setLineColor(new Color(255, 192, 0));
		}
		contentx2.setRotation(180);
		contentx2.setLineWidth(4);
		slide.addShape(contentx2);

		HSLFTextBox contentx12 = new HSLFTextBox();
		contentx12.setAnchor(new java.awt.Rectangle(755, 270, 0, 45));
		contentx12.setLineColor(new Color(0, 176, 240));
		if (Platform.equals("Instagram")) {
			contentx12.setLineColor(new Color(255, 192, 0));
		}
		contentx12.setLineWidth(4);
		slide.addShape(contentx12);

	}

	public void slideInformation(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses,
			String jumlah_data, String t1, String t2, String t3, String acc1, String acc2, String acc3)
			throws IOException {

		judul(slide, "Information", "Jumlah data, rentang waktu, dan tweet terkait");

		// ------------- add picture grafik -----------//
		HSLFPictureData pd3r = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\liker1TWv2.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3r = new HSLFPictureShape(pd3r);
		pictNew3r.setAnchor(new java.awt.Rectangle(28, 335, 75, 75));
		pictNew3r.setShapeType(ShapeType.ELLIPSE);
		slide.addShape(pictNew3r);

		// ------------- add picture grafik -----------//
		HSLFPictureData pd3r2 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\liker2TWv2.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3r2 = new HSLFPictureShape(pd3r2);
		pictNew3r2.setAnchor(new java.awt.Rectangle(250, 145, 75, 75));
		pictNew3r2.setShapeType(ShapeType.ELLIPSE);
		slide.addShape(pictNew3r2);

		// ------------- add picture grafik -----------//
		HSLFPictureData pd3r3 = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\liker3TWv2.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew3r3 = new HSLFPictureShape(pd3r3);
		pictNew3r3.setAnchor(new java.awt.Rectangle(490, 335, 75, 75));
		pictNew3r3.setShapeType(ShapeType.ELLIPSE);
		slide.addShape(pictNew3r3);

		// -------------------------Content------------------------//

		HSLFTextBox content = new HSLFTextBox();

		content.setText("Laporan ini membahas tentang topik \"" + tema + "\" dalam suatu rentang waktu " + time1
				+ " hingga " + time2);

		content.appendText(" yang beredar pada media sosial " + Platform + " dengan total data " + Kirim
				+ " yang didapat sebanyak " + jumlah_data + " data " + Post + ".", false);
		content.setAnchor(new java.awt.Rectangle(80, 440, 800, 60));
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
		content1.setAnchor(new java.awt.Rectangle(87, 245, 385, 95));
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
		content2.setAnchor(new java.awt.Rectangle(320, 60, 385, 95));
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
		content3.setAnchor(new java.awt.Rectangle(550, 245, 385, 95));
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
		content12.setAnchor(new java.awt.Rectangle(23, 410, 300, 75));

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
		content22.setAnchor(new java.awt.Rectangle(245, 220, 300, 75));

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
		content23.setAnchor(new java.awt.Rectangle(485, 410, 300, 75));

		HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
		contentP23.setAlignment(TextAlign.JUSTIFY);
		contentP23.setLineSpacing(100.0);

		HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
		runContent23.setFontSize(12.);
		runContent23.setFontFamily("Corbel");

		runContent23.setFontColor(new Color(255, 192, 0));
		slide.addShape(content23);
	}

	public void slideTweetTrend(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses)
			throws IOException {

		setBackground(slide, version);

		judul(slide, Post + " Weekly Comparison ", "Frekuensi " + Post + " perhari topik " + tema);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slideTweetTrendAll(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses)
			throws IOException {

		setBackground(slide, version);

		judul(slide, Post + " Frequency per Day", "Frekuensi " + Post + " perhari topik " + tema);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slideTweetTrenBox(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses)
			throws IOException {

		setBackground(slide, version);

		judul(slide, Post + " Distribution per Group", "Frekuensi " + Post + " perhari topik " + tema);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slideNumberofAuthorTrend(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses)
			throws IOException {

		setBackground(slide, version);

		judul(slide, "Author Weekly Comparison", "Jumlah akun yang membuat kiriman per hari");

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slideNumberofAuthorTrendAll(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2,
			String akses) throws IOException {

		setBackground(slide, version);

		judul(slide, "Author Frequency per Day", "Jumlah akun yang membuat kiriman per hari");

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slideNumberofAuthorTrendBox(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2,
			String akses) throws IOException {

		setBackground(slide, version);

		judul(slide, "Author Frequency Distribution per Group", "Jumlah akun yang membuat kiriman per hari");

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slidePostType(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses, String tb,
			String pb) throws IOException {

		setBackground(slide, version);

		judul(slide, "Type of " + Post, "Persentase tipe kiriman pada topik " + tema);
		// -------------------------KOTAK------------------------//

		HSLFTextBox content23ar = new HSLFTextBox();

		content23ar.setAnchor(new java.awt.Rectangle(565, 0, 395, 540));
		content23ar.setFillColor(new Color(204, 232, 234));
		if (Platform.equals("Instagram")) {
			content23ar.setFillColor(new Color(255, 200, 25));
		}

		slide.addShape(content23ar);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText(Post + " didominasi oleh aktivitas " + tb + " dengan persentase sebesar " + pb + ".");
		content231.setAnchor(new java.awt.Rectangle(610, 200, 300, 110));

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

	public void slidePostTypeAll(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses)
			throws IOException {

		judul(slide, "Type of " + Post + " Overall",
				"Persentase tipe kiriman pada topik " + tema + " secara keseluruhan");
		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slidePostTypeBox(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses)
			throws IOException {

		judul(slide, "Type of " + Post + " Comparison",
				"Persentase tipe kiriman pada topik " + tema + " secara keseluruhan");
		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slideTopPost(HSLFSlide slide, String time1, String time2, String akses) {

		setBackground(slide, version);

		judul(slide, "Top " + Post, Kirim + " dengan jumlah shares terbanyak");

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slideTopikBahasan(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses)
			throws IOException {

	}

	public void slideOverallNetwork(HSLFSlide slide, String time1, String time2, String akses, String a, String h,
			String t) {

		setBackground(slide, 1);

		judul(slide, "Overall Network", "Network secara keseluruhan pada topik " + tema);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText("Jumlah keseluruhan Akun: " + a + " Akun");
		content231.appendText("\nJumlah keseluruhan Hashtag: " + h + " Hashtag", false);
		content231.appendText("\nJumlah keseluruhan Tweet: " + t + " Tweet", false);
		content231.setAnchor(new java.awt.Rectangle(5, 420, 400, 10));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(14.);
		runContent231.setFontFamily("Corbel");
		runContent231.setFontColor(Color.white);
		slide.addShape(content231);

		// ========keterangan=========//
		Keterangan(slide, 1, time1, time2, akses);
	}

	public void slidehashtaguser(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses)
			throws IOException {
		setBackground(slide, version);

		judul(slide, "Top 10 Hashtag dan User",
				"Hashtag yang paling banyak digunakan dan Akun yang paling aktif membuat kiriman");

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);
	}

	public void slidehashtagtwitter(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses)
			throws IOException {
		setBackground(slide, version);

		judul(slide, "Top 10 Hashtag dan Mention",
				"Hashtag yang paling banyak digunakan dan Akun yang paling banyak disebut");

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);
	}

	public void slideTopUser(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses)
			throws IOException {
		setBackground(slide, version);

		judul(slide, "Top 10 User", "Akun yang paling aktif membuat kiriman");

		// -------------------------KOTAK------------------------//

		HSLFTextBox content23ar = new HSLFTextBox();

		content23ar.setAnchor(new java.awt.Rectangle(565, 0, 395, 540));
		content23ar.setFillColor(new Color(204, 232, 234));
		if (Platform.equals("Instagram")) {
			content23ar.setFillColor(new Color(255, 200, 25));
		}

		slide.addShape(content23ar);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);
	}

	public void slideTopUrl(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses)
			throws IOException {
		setBackground(slide, version);

		judul(slide, "Top 10 Url", "Link yang paling banyak dibagikan");

		// -------------------------KOTAK------------------------//

		HSLFTextBox content23ar = new HSLFTextBox();

		content23ar.setAnchor(new java.awt.Rectangle(565, 0, 395, 540));
		content23ar.setFillColor(new Color(204, 232, 234));
		if (Platform.equals("Instagram")) {
			content23ar.setFillColor(new Color(255, 200, 25));
		}

		slide.addShape(content23ar);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);
	}

	public void slideTop10Accounts(HSLFSlide slide, String time1, String time2, String akses) {

		setBackground(slide, version);

		judul(slide, "Top 10 Accounts", "Top Akun berdasarkan jumlah follower terbanyak dan reach terbanyak");

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

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);
	}

	public void slideTop10HashtagGrowth(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses)
			throws IOException {

		setBackground(slide, version);

		judul(slide, "Top 10 Hashtag Growth", "Hashtag dengan pertumbuhan paling besar");

		// ----- add picture-------------//
		HSLFPictureData pd = ppt.addPicture(new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\posisi.png"),
				PictureData.PictureType.PNG);
		HSLFPictureShape pictNew = new HSLFPictureShape(pd);
		pictNew.setAnchor(new java.awt.Rectangle(770, 120, 170, 80));
		slide.addShape(pictNew);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);
	}

	public void slideNetworkofRetweet(HSLFSlide slide, String time1, String time2, String akses) {

		setBackground(slide, 1);

		HSLFTextBox title = new HSLFTextBox();
		if (Platform.equals("Twitter")) {

			title.setText("Network of Retweet");
			judul(slide, "Network of Retweet", "Network " + tema + " berdasarkan aktivitas retweetnya");
		} else {
			judul(slide, "Network of Shares", "Network " + tema + " berdasarkan aktivitas sharenya");
		}

		// ========keterangan=========//
		Keterangan(slide, 1, time1, time2, akses);
	}

	public void slideNetworkofMention(HSLFSlide slide, String time1, String time2, String akses) {

		setBackground(slide, 1);

		HSLFTextBox title = new HSLFTextBox();
		if (Platform.equals("Twitter")) {

			judul(slide, "Network of Mention", "Network " + tema + " berdasarkan aktivitas mentionnya");

		} else {
			judul(slide, "Network of Post", "Network " + tema + " berdasarkan aktivitas kirimannya");
		}

		// ========keterangan=========//
		Keterangan(slide, 1, time1, time2, akses);
	}

	public void slideKesimpulan(HSLFSlide slide, String time1, String time2, String akses) {

		setBackground(slide, version);

		judul(slide, "Kesimpulan", "Analisis " + tema);

		// -------------------------Content------------------------//

		HSLFTextBox content231 = new HSLFTextBox();

		content231.setText("Berdasarkan data sampel yang disediakan oleh " + Platform + " pada tanggal " + time1
				+ " hingga " + time2 + " diakses pada " + akses + ", dengan data sebagai berikut:");
		content231.setAnchor(new java.awt.Rectangle(30, 90, 900, 110));

		HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
		contentP231.setAlignment(TextAlign.JUSTIFY);
		contentP231.setLineSpacing(100.0);

		HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
		runContent231.setFontSize(18.);
		runContent231.setFontFamily("Calibri");

		slide.addShape(content231);

	}

}
