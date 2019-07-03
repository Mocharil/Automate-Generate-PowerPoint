package com.ebdesk.report;

import com.ebdesk.report.slide.SNAInstagramReport;
import com.ebdesk.report.slide.SNAInstagramTopicv2;
import com.ebdesk.report.slide.SNATelegramChannel;
import com.ebdesk.report.slide.SNATelegramGroup;
import com.ebdesk.report.slide.SNATwitter9Topic;
import com.ebdesk.report.slide.SNATwitterAccount;
import com.ebdesk.report.slide.SNATwitterAccountv2;

import java.awt.Color;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

import org.bson.Document;

import com.mongodb.MongoClient;
import com.mongodb.MongoCredential;
import com.mongodb.ServerAddress;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;

import com.mongodb.spark.rdd.api.java.JavaMongoRDD;

import com.ebdesk.report.slide.SNATwitterReport;
import com.ebdesk.report.slide.SNATwitterTopic;
import com.ebdesk.report.slide.SNATwitterTopicv2;
import com.ebdesk.report.slide.SNATwitterTopicv3;
import com.ebdesk.report.slide.SNAWhatsApp1Group;
import com.ebdesk.report.slide.TimeSeriesChart;
import com.ebdesk.report.slide.Wordcloud;
import com.ebdesk.report.slide.backupv1;
import com.ebdesk.report.slide.backupv2;
import com.ebdesk.report.slide.backupv3;
import com.ebdesk.report.slide.backupv4;
import com.ebdesk.report.slide.grafik;
import com.ebdesk.report.slide.percobaan;
import com.ebdesk.report.slide.readjson;
import com.kennycason.kumo.WordFrequency;
import com.ebdesk.report.slide.GrafikDonat;
import com.ebdesk.report.slide.PieChartReport;
import com.ebdesk.report.slide.SNAAllPlatform;
import com.ebdesk.report.slide.SNAAllPlatformMalay;
import com.ebdesk.report.slide.SNAAllPlatformv2;
import com.ebdesk.report.slide.SNAAllPlatformv3;
import com.ebdesk.report.slide.SNAAllPlatformv4;
import com.ebdesk.report.slide.SNAFacebookReport;
import com.ebdesk.report.slide.SNAFacebookTopicv2;
import com.ebdesk.report.slide.SNAFacebookTopicv3;
import com.ebdesk.report.slide.SNAInstagramAccount;

import javafx.stage.Stage;

// 960 x 540
// 720 x 540

public class App {
	public static void main(String[] args) throws Exception {
		System.out.println("mulai");
//		 System.out.println("Argument one = "+args[0]);
//		 System.out.println("Argument two = "+args[1]);
//		 System.out.println("Argument three = "+args[2]);

		
		 SNAAllPlatformMalay report1 = new SNAAllPlatformMalay( 2, 960, 540, "Analisis 9 Tokoh Politik", "Twitter", args );
		 report1.execute();
		
//		 SNAAllPlatform report = new SNAAllPlatform( 2, 960, 540, "Analisis 9 Tokoh Politik", "Twitter");
//		 report.execute();
		
//		 SNAAllPlatformv2 report1 = new SNAAllPlatformv2( 2, 960, 540, "Analisis 9 Tokoh Politik", "Twitter");
//		 report1.execute();

		// SNAAllPlatformv3 report1 = new SNAAllPlatformv3( 2, 960, 540, "Analisis 9
		// Tokoh Politik","Twitter");
		// report1.execute();

		// SNAAllPlatformv4 report1 = new SNAAllPlatformv4( 2, 960, 540, "Analisis 9
		// Tokoh Politik","Twitter");
		// report1.execute();
//		if (args[2].equals("1")) {
//			backupv1 report1 = new backupv1(2, 960, 540, args[0], args[1]);
//			report1.execute();
//		} else if (args[2].equals("2")) {
//			backupv2 report2 = new backupv2(2, 960, 540, args[0], args[1]);
//			report2.execute();
//		} else if (args[2].equals("3")) {
//			backupv3 report3 = new backupv3(2, 960, 540, args[0], args[1]);
//			report3.execute();
//		} else if (args[2].equals("4")) {
//			backupv4 report4 = new backupv4(2, 960, 540, args[0], args[1]);
//			report4.execute();
//		} else {
//			backupv1 report1 = new backupv1(2, 960, 540, args[0], args[1]);
//			report1.execute();
//		}
		// ======================INSTAGRAM======================//

		// SNAInstagramAccount report = new SNAInstagramAccount
		// ("slideshowInstagramAccount.ppt", 3, 960, 540, "Pilgub Kaltim Paslon 4");
		// report.execute();

		// ======================GRAFIK======================//

		// TimeSeriesChart chart = new TimeSeriesChart();
		// chart.TimeSeries(null);

		// String[] headers = { "Count", "Follower", "Username", "Reach"};
		// table table = new table();
		// table.table("table21.png", headers);}

		// grafik chart1 = new grafik();
		// chart1.mulai(null, "Tweet", 10, "Reply", 120, "Quoted", 34, "Retweet", 55,
		// "pie.png", 4, "rgba(84,130,53 )");

		// PieChartReport chart1 = new PieChartReport();
		// chart1.pie("Tweet", 10, "Reply", 120, "Quoted", 34, "Retweet", 55, "pie.png",
		// 4);

		// PieChartReport chart1 = new PieChartReport();
		// chart1.pie("Tweet", 10, "Reply", 120, "Quoted", 34, "Retweet", 55, "pie.png",
		// Color.white);

		// List<String> myString = new ArrayList<String>();
		// myString.add("ABS");
		// Wordcloud chart = new Wordcloud();
		// List<WordFrequency> x = chart.readCNN("wordcloudTeleChannel.png", myString,
		// "orange");
		//

		// ======================FACEBOOK======================//

		// ======================TWITTER======================//

		// SNATwitterAccountv2 report = new
		// SNATwitterAccountv2("slideTwitterAccountv2.ppt", 3, 960, 540,
		// "@ShamsiAli2");
		// report.execute();

		// ======================TELEGRAM======================//

		// SNATelegramChannel report = new
		// SNATelegramChannel("slideshowTeleChannel.ppt", 2, 960, 540,
		// "Pilgub Kaltim Paslon 4");
		// report.execute();

		// SNATelegramGroup report = new SNATelegramGroup ("slideshowTeleGroup.ppt", 2,
		// 960, 540, "Pilgub Kaltim Paslon 4");
		// report.execute();

		// ======================WHATSAPP======================//

		// SNAWhatsApp1Group report = new
		// SNAWhatsApp1Group("slideshowWhatsApp1Group.ppt", 3, 960, 540, "Pilgub Kaltim
		// Paslon 4");
		// report.execute();

		////////////////////// ==================================SELESAI================================////////////////////

		// ----------------------------IG

		// SNAInstagramTopicv2 report = new
		// SNAInstagramTopicv2("slideshowIGtopicv2.ppt", 2, 960, 540,
		// "Pilgub Kaltim Paslon 4", "criteria_instagram",
		// "330a59fae963a979e2d96a9ec91e8b7c");
		// report.execute();

		// SNAInstagramReport reportIG = new
		// SNAInstagramReport("slideshowIGTopicv1.ppt", 2, 960, 540, "Pilgub Kaltim
		// Paslon 4", "criteria_instagram", "330a59fae963a979e2d96a9ec91e8b7c");
		// reportIG.execute();

		// ----------------------------FB

		// SNAFacebookReport reportFB = new SNAFacebookReport("slideshowFBTopicv1.ppt",
		// 1, 960, 540,
		// "Pilgub Kaltim Paslon 4", "criteria_facebook",
		// "330a59fae963a979e2d96a9ec91e8b7c");
		// reportFB.execute();

		// SNAFacebookTopicv2 report = new SNAFacebookTopicv2("slideshowFBtopicv2.ppt",
		// 2, 960, 540,
		// "Pilgub Kaltim Paslon 4", "criteria_facebook",
		// "330a59fae963a979e2d96a9ec91e8b7c");
		// report.execute();

		// SNAFacebookTopicv3 report = new SNAFacebookTopicv3("slideshowFBtopicv3.ppt",
		// 2, 960, 540, "FreePort");
		// report.execute();

		// -----------------------------TW

		// SNATwitterAccount report = new
		// SNATwitterAccount("slideshowTwitterAccountv1.ppt", 2, 960, 540,
		// "ShamsiAli2", "10", "ShamsiAli2");
		// report.execute();

		// SNATwitterTopicv2 report = new SNATwitterTopicv2("slideshowTWtopicv2.ppt", 2,
		// 960, 540,
		// "Pilkada Kaltim Paslon 4", "criteria_twitter",
		// "ff7b81e29124d2623c36bd1102d6cc1a");
		// report.execute();

		// SNATwitterReport report = new SNATwitterReport("slideshowTW.ppt", 1, 960,
		// 540, "Pilkada Kaltim Paslon 4",
		// "criteria_twitter", "330a59fae963a979e2d96a9ec91e8b7c");
		// report.execute();

		// SNATwitterTopicv3 report = new SNATwitterTopicv3("slideshowTWtopicv3.ppt", 2,
		// 960, 540, "PRO/KONTRA JOKOWI");
		// report.execute();

		// SNATwitter9Topic report = new SNATwitter9Topic("slideshowTW9Topic.ppt", 2,
		// 960, 540, "PRO/KONTRA JOKOWI");
		// report.execute();

		// readjson report = new readjson("slideshowTWtopicv3.ppt", 2, 960, 540,
		// "PRO/KONTRA JOKOWI ", "criteria_twitter",
		// "ff7b81e29124d2623c36bd1102d6cc1a");
		// report.execute();
	}

}