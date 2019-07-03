package com.ebdesk.report.slide;

import java.awt.Color;
import java.io.File;
import java.io.IOException;

import javax.imageio.ImageIO;

import org.jfree.chart.ChartUtilities;
import org.jfree.chart.JFreeChart;
import org.jfree.chart.plot.CategoryPlot;
import org.jfree.chart.renderer.category.BarRenderer;
import org.jfree.chart.renderer.category.StandardBarPainter;

import javafx.application.Application;
import javafx.collections.*;
import javafx.embed.swing.SwingFXUtils;
import javafx.scene.Scene;
import javafx.scene.SnapshotParameters;
import javafx.scene.chart.PieChart;
import javafx.scene.image.WritableImage;
import javafx.scene.layout.StackPane;
import javafx.stage.Stage;

import javafx.application.Application;
import javafx.scene.Group;
import javafx.scene.Scene;

import javafx.scene.shape.Circle;
import javafx.scene.text.Font;
import javafx.scene.text.Text;
import javafx.stage.Stage;

public class GrafikDonat extends Application {

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
		this.rgba = rgba;
		this.tweet = tweet;
		this.reply_ = reply_;
		this.reply = reply;
		this.quoted_ = quoted_;
		this.quoted = quoted;
		this.retweet_ = retweet_;
		this.retweet = retweet;
		this.FILE = FILE;
		this.versi = versi;
		launch(args);
	}

	@Override
	public void start(Stage stage) throws IOException {
		stage.setWidth(700);
		stage.setHeight(700);
		
		if (versi == 4) {
			ObservableList<PieChart.Data> pieChartData = FXCollections.observableArrayList(
					new PieChart.Data(tweet_ + " : " + tweet + "%", tweet),
					new PieChart.Data(reply_ + " : " + reply + "%", reply),
					new PieChart.Data(quoted_ + " : " + quoted + "%", quoted),
					new PieChart.Data(retweet_ + " : " + retweet + "%", retweet));
			save( pieChartData,  stage );
		}

		else {

			ObservableList<PieChart.Data> pieChartData = FXCollections.observableArrayList(
					new PieChart.Data(tweet_ + " : " + tweet + "%", tweet),
					new PieChart.Data(reply_ + " : " + reply + "%", reply),
					new PieChart.Data(quoted_ + " : " + quoted + "%", quoted));
			save( pieChartData,  stage );
		}

	}

	private static void save(ObservableList<PieChart.Data> pieChartData, Stage stage) throws IOException {
		final DoughnutChart chart = new DoughnutChart(pieChartData);
		chart.setLegendVisible(false);
		chart.setStyle("-fx-background-color: " + rgba + ";");
		Scene scene = new Scene(new StackPane(chart));
		stage.setScene(scene);
		stage.show();
		// save//
		WritableImage image = chart.snapshot(new SnapshotParameters(), null);

		File file = new File("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\" + FILE);

		try {
			ImageIO.write(SwingFXUtils.fromFXImage(image, null), "png", file);
			System.out.println("berhasil");
			stage.close();

		} catch (IOException e) {
			System.out.println("gagal");
		}

	}

}