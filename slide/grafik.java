package com.ebdesk.report.slide;

import java.awt.Color;
import java.awt.Insets;
import java.awt.Label;
import java.awt.image.RenderedImage;
import java.io.File;
import java.io.IOException;
import javafx.embed.swing.SwingFXUtils;
import javax.imageio.ImageIO;

import org.jfree.chart.ChartFactory;
import org.jfree.chart.ChartFrame;
import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;

import javafx.application.Application;
import javafx.geometry.Side;
import javafx.scene.Scene;
import javafx.scene.SnapshotParameters;
import javafx.scene.chart.PieChart;
import javafx.scene.image.WritableImage;
import javafx.scene.layout.StackPane;
import javafx.stage.Stage;
import javafx.stage.StageStyle;
import javafx.application.Application;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Paint;
import javafx.scene.Group;
import javafx.scene.Scene;
import javafx.scene.shape.Circle;
import javafx.scene.text.Font;
import javafx.scene.text.Text;
import javafx.stage.Stage;

public class grafik extends Application {
	private static String tweet_;
	private static float tweet;
	private static String reply_;
	private static float reply;
	private static String quoted_;
	private static float quoted;
	private static String retweet_;
	private static float retweet;
	private static String FILE;
	private static int versi;
	private static String rgba;

	public void mulai(String[] args, String tweet_, float tweet, String reply_, float reply, String quoted_,
			float quoted, String retweet_, float retweet, String FILE, int versi, String rgba) throws Exception {
		this.tweet_ = tweet_;
		this.tweet = tweet;
		this.reply_ = reply_;
		this.reply = reply;
		this.quoted_ = quoted_;
		this.quoted = quoted;
		this.retweet_ = retweet_;
		this.retweet = retweet;
		this.FILE = FILE;
		this.versi = versi;
		this.rgba= rgba;
		launch(args);
	}

	public void start(Stage primaryStage) throws Exception {

		PieChart pieChart = new PieChart();


		if (versi == 4) {

			PieChart.Data slice1 = new PieChart.Data(tweet_ + " : " + tweet + "%", tweet);
			PieChart.Data slice2 = new PieChart.Data(reply_ + " : " + reply + "%", reply);
			PieChart.Data slice3 = new PieChart.Data(quoted_ + " : " + quoted + "%", quoted);
			PieChart.Data slice4 = new PieChart.Data(retweet_ + " : " + retweet + "%", retweet);
			pieChart.getData().add(slice1);
			pieChart.getData().add(slice2);
			pieChart.getData().add(slice3);
			pieChart.getData().add(slice4);
		}
		if (versi == 3) {
			PieChart.Data slice1 = new PieChart.Data(tweet_ + " : " + tweet + "%", tweet);
			PieChart.Data slice2 = new PieChart.Data(reply_ + " : " + reply + "%", reply);
			PieChart.Data slice3 = new PieChart.Data(quoted_ + " : " + quoted + "%", quoted);
			pieChart.getData().add(slice1);
			pieChart.getData().add(slice2);
			pieChart.getData().add(slice3);
			

		}
		pieChart.setLegendVisible(false);
	
		
		StackPane root = new StackPane(pieChart);
		Scene scene = new Scene(root, 780, 800 );
		root.setStyle("-fx-background-color: "+ rgba +";");
		
		primaryStage.setScene(scene);
		primaryStage.show();
		SnapshotParameters parameters = new SnapshotParameters();

		WritableImage image = root.snapshot(parameters, null);
		File file = new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\" + FILE);

		try {
			ImageIO.write(SwingFXUtils.fromFXImage(image, null), "png", file);
			System.out.println("berhasil");
			primaryStage.close();
			
		} catch (IOException e) {
			System.out.println("gagal");
		}

	}

}