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

public class SNAAllPlatformMalay {

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
	private String[] args;

	public SNAAllPlatformMalay(int version, int width, int height, String tema, String Platform, String[] args) {
		this.filename = Platform + "_" + tema + "_Malay.ppt";
		this.version = version;
		this.height = height;
		this.width = width;
		this.tema = tema;
		this.args = args;
		this.Platform = Platform;
		if (Platform.equals("Twitter")) {
			this.Post = "Tweet";
			this.Kirim = "Tweet";
		} else {
			this.Post = "Post";
			this.Kirim = "Post";
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
		List<String> listnamatop = new ArrayList<String>();

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
							.replace("post types percentages", "").replace("tokoh", ""));
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
							.replace("postweeklycomparison", "").replace("tokoh", ""));
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
							.replace("authorweeklycomparison", "").replace("tokoh", ""));
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
							.replace("topsharedurl", "").replace("tokoh", ""));
				} catch (Exception e) {

				}
			}

			if (String.valueOf(file.getCanonicalPath()).contains("topusedhashtag") == true) {
				hashtag.add(String.valueOf(file.getCanonicalPath()));
				try {
					listnamahashtag.add(String.valueOf(file.getName()).replace("_", "").replace(".png", "")
							.replace("topusedhashtag", "").replace("tokoh", ""));
				} catch (Exception e) {

				}
			}
			if (String.valueOf(file.getCanonicalPath()).contains("topmentioneduser") == true) { // belum
				mention.add(String.valueOf(file.getCanonicalPath()));
				try {
					listnamamention.add(String.valueOf(file.getName()).replace("_", "").replace(".png", "")
							.replace("topmentioneduser", "").replace("tokoh", ""));
				} catch (Exception e) {

				}
			}
			if (String.valueOf(file.getCanonicalPath()).contains("topactiveuser") == true) {
				user.add(String.valueOf(file.getCanonicalPath()));
				try {
					listnamauser.add(String.valueOf(file.getName()).replace("_", "").replace(".png", "")
							.replace("topactiveuser", "").replace("tokoh", ""));
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
							.replace("top post share", "").replace("tokoh", ""));
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
							.replace("wordcloud", "").replace("tokoh", ""));
				} catch (Exception e) {

				}

			}
			if (String.valueOf(file.getCanonicalPath()).contains("hashtagcloud") == true) {

				wchtg.add(String.valueOf(file.getCanonicalPath()));
				try {
					listnamahcloud.add(String.valueOf(file.getName()).replace("_", "").replace(".png", "")
							.replace("hashtagcloud", "").replace("tokoh", ""));
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

		while (scanner.hasNext()) {

			List<String> line = parseLine(scanner.nextLine());
			try {

				map.put(line.get(0), line.get(1));
				line0.add(line.get(0));

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
		slideMaster(ppt, slide1, time1, time2, akses);
		slide1.addShape(pictNewl);

		// slide 3
		HSLFSlide slide21 = ppt.createSlide();
		try {
			slidePreliminary(ppt, slide21, time1, time2, akses, jumlah_data);
			slide21.addShape(pictNewl);
		} catch (Exception e) {
		}

		// slide Tweet Frequency
		HSLFSlide slide5 = ppt.createSlide();
		try {
			slideTweetTrend(ppt, slide5, time1, time2, akses, jumlah_data, jumlah_data);
			slide5.addShape(pictNewl);
		} catch (Exception e) {
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
		}

		for (int k = 0; k < list.size(); k++) {
			for (int i = 0; i < list.size(); i++) {
				try {
					if (wc.get(i).contains(ww.get(k)) || wchtg.get(i).contains(ww.get(k))) {

						setBackground(topbahasan.get(k), version);

						judul(topbahasan.get(k), "Wordcloud Topic \"" + tema + "\"", "Most used/appeared word");

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

						runContent23.setFontColor(new Color(255, 192, 0));
						topbahasan.get(k).addShape(content23);
						topbahasan.get(k).addShape(pictNewl);

						try {
							if (wc.get(i).contains(ww.get(k))) {

								// ----- add picture-------------//
								HSLFPictureData pd = ppt.addPicture(new File(wc.get(i)), PictureData.PictureType.PNG);
								HSLFPictureShape pictNew = new HSLFPictureShape(pd);
								pictNew.setAnchor(new java.awt.Rectangle(20, 90, 410, 410));
								topbahasan.get(k).addShape(pictNew);

								// -------------------------Content------------------------//
								int h = 0;
								HSLFTextBox content231 = new HSLFTextBox();
								for (int x = 0; x < map.size(); x++) {

									try {
										if (line0.get(x).contains("most_word_1_") && line0.get(x).contains(ww.get(k))) {

											content231.setText("The words mostly used in topic " + tema + " are ");
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
												content231.appendText(", ", false);
											} else {
												content231.setText("The words mostly used in topic " + tema + " are ");
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
												content231.appendText(", ", false);
											}

											else if (h == 5) {
												content231.appendText(" and ", false);

											} else {
												content231.setText("The words mostly used in topic " + tema + " are ");
											}
											content231.appendText(map.get(line0.get(z2)) + ".", false);
										}
									} catch (Exception e) {
									}
								}

								content231.setAnchor(new java.awt.Rectangle(450, 135, 400, 110));
								HSLFTextParagraph contentP231 = content231.getTextParagraphs().get(0);
								contentP231.setAlignment(TextAlign.JUSTIFY);
								contentP231.setLineSpacing(100.0);

								HSLFTextRun runContent231 = contentP231.getTextRuns().get(0);
								runContent231.setFontSize(20.);
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
		List<HSLFSlide> slide21x = new ArrayList<HSLFSlide>();
		for (int i = 0; i < 3; i++) {
			slide21x.add(ppt.createSlide());
			slideTweetsDay(slide21x.get(i), time1, time2, akses);
			slide21x.get(i).addShape(pictNewl);
		}
		// Slide Network of Retweet or Shares
		List<HSLFSlide> slide19 = new ArrayList<HSLFSlide>();
		for (int i = 0; i < 2; i++) {
			slide19.add(ppt.createSlide());
			slideNetwork(slide19.get(i), time1, time2, akses);
			slide19.get(i).addShape(pictNewl);
		}

		// slide Top Tweet/Post
		List<HSLFSlide> top10 = new ArrayList<HSLFSlide>();
		for (int i = 0; i < top10tweet.size(); i++) {
			top10.add(ppt.createSlide());
			judul(top10.get(i), "Top 5 Tweet", "Tweet with most Favorite and Retweet");

			top10.get(i).addShape(top10tweet.get(i));
			top10tweet.get(i).moveTo(10, 100);

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

		// Slide Top 10 Hashtag and Url

		List<HSLFSlide> slidehashtag = new ArrayList<HSLFSlide>();
		List<String> jj = new ArrayList<String>();

		HashSet<String> sethashtagmention = new HashSet<String>(listnamaurl);
		sethashtagmention.addAll(listnamahashtag);
		jj.addAll(sethashtagmention);
		Collections.sort(jj, Collections.reverseOrder());
		for (int k = 0; k < sethashtagmention.size(); k++) {
			slidehashtag.add(ppt.createSlide());
		}

		for (int k = 0; k < list.size(); k++) {
			for (int i = 0; i < list.size(); i++) {
				try {
					if (Platform.equals("Twitter")) {

						if (hashtag.get(i).contains(jj.get(k)) || url.get(i).contains(jj.get(k))) {
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
								if (url.get(i).contains(jj.get(k))) {
									// ----- add picture-------------//
									HSLFPictureData pd1 = ppt.addPicture(new File(url.get(i)),
											PictureData.PictureType.PNG);
									HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
									pictNew1.setAnchor(new java.awt.Rectangle(500, 80, 420, 300));
									slidehashtag.get(k).addShape(pictNew1);

									// -------------------------Content------------------------//
									int h = 0;
									HSLFTextBox content231 = new HSLFTextBox();
									for (int x = 0; x < map.size(); x++) {

										try {
											if (line0.get(x).contains("most_url_1_")
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
											if (line0.get(z1).contains("most_url_2_")
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
											if (line0.get(z2).contains("most_url_3_")
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

				} catch (Exception e) {
				}

			}
		}

		// Slide Top 10 User and Mention

		List<HSLFSlide> slideuser = new ArrayList<HSLFSlide>();
		List<String> kk = new ArrayList<String>();

		HashSet<String> setusermention = new HashSet<String>(listnamauser);
		sethashtagmention.addAll(listnamamention);
		kk.addAll(setusermention);
		Collections.sort(kk, Collections.reverseOrder());
		for (int k = 0; k < setusermention.size(); k++) {
			slideuser.add(ppt.createSlide());
		}

		for (int k = 0; k < list.size(); k++) {
			for (int i = 0; i < list.size(); i++) {
				try {
					if (Platform.equals("Twitter")) {

						if (user.get(i).contains(kk.get(k)) || mention.get(i).contains(kk.get(k))) {
							slideusertwitter(ppt, slideuser.get(k), time1, time2, akses);

							HSLFTextBox content23 = new HSLFTextBox();
							content23.setText(kk.get(k));

							content23.setAnchor(new java.awt.Rectangle(50, 20, 860, 75));

							HSLFTextParagraph contentP23 = content23.getTextParagraphs().get(0);
							contentP23.setAlignment(TextAlign.RIGHT);
							contentP23.setLineSpacing(100.0);

							HSLFTextRun runContent23 = contentP23.getTextRuns().get(0);
							runContent23.setFontSize(28.);
							runContent23.setItalic(true);
							runContent23.setFontFamily("Corbel");

							runContent23.setFontColor(new Color(255, 192, 0));
							slideuser.get(k).addShape(content23);
							slideuser.get(k).addShape(pictNewl);
							try {
								if (user.get(i).contains(kk.get(k))) {

									// ----- add picture-------------//
									HSLFPictureData pd = ppt.addPicture(new File(user.get(i)),
											PictureData.PictureType.PNG);
									HSLFPictureShape pictNew = new HSLFPictureShape(pd);
									pictNew.setAnchor(new java.awt.Rectangle(40, 80, 420, 300));
									slideuser.get(k).addShape(pictNew);

									// -------------------------Content------------------------//
									int h = 0;
									HSLFTextBox content231 = new HSLFTextBox();
									for (int x = 0; x < map.size(); x++) {

										try {
											if (line0.get(x).contains("most_active_1_")
													&& line0.get(x).contains(kk.get(k))) {

												content231.setText("Hashtag yang paling banyak digunakan pada topik "
														+ tema + " adalah ");
												content231.appendText("@" + map.get(line0.get(x)), false);
												h = 5;

											}
										} catch (Exception e) {

										}

									}

									for (int z1 = 0; z1 < map.size(); z1++) {
										try {
											if (line0.get(z1).contains("most_active_2_")
													&& line0.get(z1).contains(kk.get(k))) {

												if (h == 5) {
													content231.appendText(". Disusul oleh hashtag ", false);
												} else {
													content231.setText("Hashtag yang paling aktif pada topik " + tema
															+ " adalah ");
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
													&& line0.get(z2).contains(kk.get(k))) {

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
												content231.appendText("@" + map.get(line0.get(z2)) + ".", false);
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

									slideuser.get(k).addShape(content231);

								}
							} catch (Exception e) {
							}
							try {
								if (mention.get(i).contains(kk.get(k))) {
									// ----- add picture-------------//
									HSLFPictureData pd1 = ppt.addPicture(new File(mention.get(i)),
											PictureData.PictureType.PNG);
									HSLFPictureShape pictNew1 = new HSLFPictureShape(pd1);
									pictNew1.setAnchor(new java.awt.Rectangle(500, 80, 420, 300));
									slideuser.get(k).addShape(pictNew1);

									// -------------------------Content------------------------//
									int h = 0;
									HSLFTextBox content231 = new HSLFTextBox();
									for (int x = 0; x < map.size(); x++) {

										try {
											if (line0.get(x).contains("most_mention_1_")
													&& line0.get(x).contains(kk.get(k))) {

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
													&& line0.get(z1).contains(kk.get(k))) {

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
													&& line0.get(z2).contains(kk.get(k))) {

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

									slideuser.get(k).addShape(content231);

								}
							} catch (Exception e) {
							}
						}

					}

				} catch (Exception e) {
				}
			}
		}

		// Slide Kesimpulan
		HSLFSlide slide22 = ppt.createSlide();
		slideKesimpulan(slide22, time1, time2, akses);
		slide22.addShape(pictNewl);

		// Slide Kesimpulan
		HSLFSlide slide22f = ppt.createSlide();
		setBackground(slide22f, 1);

		// -------------------------Content------------------------//

		HSLFTextBox content23a = new HSLFTextBox();

		content23a.setText("Attachment");

		content23a.setAnchor(new java.awt.Rectangle(50, 170, 860, 50));

		HSLFTextParagraph contentP23a = content23a.getTextParagraphs().get(0);
		contentP23a.setAlignment(TextAlign.CENTER);
		contentP23a.setLineSpacing(100.0);

		HSLFTextRun runContent23a = contentP23a.getTextRuns().get(0);
		runContent23a.setFontSize(44.);
		runContent23a.setFontFamily("Calibri Light (Headings)");
		runContent23a.setFontColor(Color.WHITE);

		slide22f.addShape(content23a);
		// -------------------------Content------------------------//

		HSLFTextBox content23ar = new HSLFTextBox();

		content23ar.setText("Profiling of Influencer Accounts Related to Topic");

		content23ar.setAnchor(new java.awt.Rectangle(50, 240, 860, 50));

		HSLFTextParagraph contentP23ar = content23ar.getTextParagraphs().get(0);
		contentP23ar.setAlignment(TextAlign.CENTER);
		contentP23ar.setLineSpacing(100.0);

		HSLFTextRun runContent23ar = contentP23ar.getTextRuns().get(0);
		runContent23ar.setFontSize(18.);
		runContent23ar.setFontFamily("Calibri (Body)");
		runContent23ar.setFontColor(Color.WHITE);
		slide22f.addShape(content23ar);


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
		if (atas.contains("Network") || bawah.contains("Network")) {
			runIssue.setFontColor(Color.white);
		} else {
			runIssue.setFontColor(Color.black);
		}
		slide.addShape(issue);

		// =========garis===========//
		Garis(slide, 30, 78, 100, 1);

		// =======content Platform========//
		Platform(slide);

	}

	public void Keterangan(HSLFSlide slide, int mode, String time1, String time2, String akses) {

		HSLFTextBox contentp = new HSLFTextBox();

		contentp.setText("Ket: Berdasarkan data sampel yang disediakan oleh " + Platform + " pada tanggal " + time1
				+ " sampai " + time2 + ", diakses pada " + akses + ".");

		contentp.setAnchor(new java.awt.Rectangle(250, 510, 700, 35));

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

	public void slideMaster(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses)
			throws IOException {
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
		runTitle.setFontColor(new Color(61, 140, 255));
		if (Platform.equals("Instagram")) {
			runTitle.setFontColor(new Color(255, 192, 0));
		}

		slide.addShape(title);

		// --------------------Investigation--------------------
		HSLFTextBox issue = new HSLFTextBox();
		issue.setText("Topic Analysis" + tema + " : " + time1 + " - " + time2);
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

		// --------------------Investigation--------------------
		HSLFTextBox issue1 = new HSLFTextBox();
		issue1.setText(akses);
		issue1.setAnchor(new java.awt.Rectangle(5, 520, 945, 50));
		HSLFTextParagraph issueP1 = issue1.getTextParagraphs().get(0);
		issueP1.setAlignment(TextAlign.CENTER);
		issueP1.setSpaceAfter(0.);
		issueP1.setSpaceBefore(0.);

		HSLFTextRun runIssue1 = issueP1.getTextRuns().get(0);
		runIssue1.setFontSize(14.);
		runIssue1.setFontFamily("Calibri");
		runIssue1.setBold(false);
		runIssue1.setFontColor(new Color(166, 166, 166));
		slide.addShape(issue1);

		// =========garis===========//
		Garis(slide, 430, 308, 100, 1);

		// =======content Platform========//
		Platform(slide);

	}

	public void slidePreliminary(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses,
			String jumlah_data) throws IOException {

		judul(slide, "Preliminary", "Total Data, Time Range, and Keywords");

		// -------------------------Content------------------------//

		HSLFTextBox content = new HSLFTextBox();

		content.setText("1. Keyword for " + tema + Platform + "analysis are:");
		try {
			for (int i = 0; i < args.length; i++) {
				content.appendText("\n  • " + args[i], false);
			}
		} catch (Exception e) {
		}
		content.appendText("\n\n2. Data being crawled in " + akses + ", from " + time1 + " until " + time2
				+ ". The amount of data is obtained about " + jumlah_data + " " + Post + "s.", false);
		content.setAnchor(new java.awt.Rectangle(25, 100, 800, 50));

		HSLFTextParagraph contentP = content.getTextParagraphs().get(0);
		contentP.setAlignment(TextAlign.LEFT);

		contentP.setLineSpacing(100.0);

		HSLFTextRun runContent = contentP.getTextRuns().get(0);
		runContent.setFontSize(17.5);
		runContent.setFontFamily("Calibri");

		slide.addShape(content);

	}

	public void slideTweetTrend(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses, String tb,
			String jumlah_data) throws IOException {

		setBackground(slide, version);

		judul(slide, Post + " Frequency Topic \"" + tema + "\"", "");

		// -------------------------Content------------------------//

		HSLFTextBox content23a = new HSLFTextBox();

		content23a.setText("The development of issues about \"" + tema + "\" was mostly done through ");
		content23a.appendText(tb + " activities.", false);
		content23a.setAnchor(new java.awt.Rectangle(70, 360, 320, 50));

		HSLFTextParagraph contentP23a = content23a.getTextParagraphs().get(0);
		contentP23a.setAlignment(TextAlign.JUSTIFY);
		contentP23a.setLineSpacing(100.0);

		HSLFTextRun runContent23a = contentP23a.getTextRuns().get(0);
		runContent23a.setFontSize(16.);
		runContent23a.setFontFamily("Calibri (Body)");

		slide.addShape(content23a);

		// -------------------------Content------------------------//

		HSLFTextBox content23ax = new HSLFTextBox();

		content23ax.setText("The amount of tweets that discussed about \"" + tema + "\" from " + time1 + " until "
				+ time2 + " is " + jumlah_data + " " + Post + "s.");
		content23ax.setAnchor(new java.awt.Rectangle(450, 320, 460, 50));

		HSLFTextParagraph contentP23ax = content23ax.getTextParagraphs().get(0);
		contentP23ax.setAlignment(TextAlign.JUSTIFY);
		contentP23ax.setLineSpacing(100.0);

		HSLFTextRun runContent23ax = contentP23ax.getTextRuns().get(0);
		runContent23ax.setFontSize(16.);
		runContent23ax.setFontFamily("Calibri (Body)");

		slide.addShape(content23ax);

	}

	public void slideTweetsDay(HSLFSlide slide, String time1, String time2, String akses) {

		setBackground(slide, 1);

		judul(slide, Post + "s/Day About " + tema,
				"Network formed by retweet, mention, reply, quoted and using hashtag activity related to " + tema
						+ " issue.");

		// ========keterangan=========//

	}

	public void slideTopPost(HSLFSlide slide, String time1, String time2, String akses) {

		setBackground(slide, version);

		// ========keterangan=========//
		Keterangan(slide, 2, time1, time2, akses);

	}

	public void slidehashtaguser(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses)
			throws IOException {
		setBackground(slide, version);

		judul(slide, "Top 10 hashtag dan User",
				"Hashtag yang paling banyak digunakan dan Akun yang paling aktif membuat kiriman");

		// ========keterangan=========//

	}

	public void slidehashtagtwitter(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses)
			throws IOException {
		setBackground(slide, version);

		judul(slide, "Top 10 Hashtag and Url",
				"Hashtag yang paling banyak digunakan dan Akun yang paling banyak disebut");

		// ========keterangan=========//

	}

	public void slideusertwitter(HSLFSlideShow ppt, HSLFSlide slide, String time1, String time2, String akses)
			throws IOException {
		setBackground(slide, version);

		judul(slide, "Top 10 User and Mention",
				"Hashtag yang paling banyak digunakan dan Akun yang paling banyak disebut");

		// ========keterangan=========//

	}

	public void slideNetwork(HSLFSlide slide, String time1, String time2, String akses) {

		setBackground(slide, 1);

		HSLFTextBox title = new HSLFTextBox();

		judul(slide, "Network about " + tema,
				"Network formed by retweet, mention, reply, quoted and using hashtag activity related to " + tema
						+ " issue");

		// ========keterangan=========//

	}

	public void slideKesimpulan(HSLFSlide slide, String time1, String time2, String akses) {

		setBackground(slide, version);

		judul(slide, "Conclusions", "Topic \"" + tema + "\"");

	}

}
