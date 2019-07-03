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
		// example
		 SNAAllPlatformMalay report1 = new SNAAllPlatformMalay( 2, 960, 540, "Analisis 9 Tokoh Politik", "Twitter", args );
		 report1.execute();

	}

}
