package com.ebdesk.report.slide;

import com.kennycason.kumo.CollisionMode;
import com.kennycason.kumo.WordCloud;
import com.kennycason.kumo.WordFrequency;
import com.kennycason.kumo.bg.CircleBackground;
import com.kennycason.kumo.bg.PixelBoundryBackground;
import com.kennycason.kumo.bg.RectangleBackground;
import com.kennycason.kumo.font.FontWeight;
import com.kennycason.kumo.font.KumoFont;
import com.kennycason.kumo.font.scale.LinearFontScalar;
import com.kennycason.kumo.font.scale.SqrtFontScalar;
import com.kennycason.kumo.image.AngleGenerator;
import com.kennycason.kumo.nlp.FrequencyAnalyzer;
import com.kennycason.kumo.palette.ColorPalette;
import org.apache.commons.io.IOUtils;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.List;

/**
 * Created by kenny on 6/29/14.
 */
public class Wordcloud {

	private static final Logger LOGGER = LoggerFactory.getLogger(Wordcloud.class);

	private static final Random RANDOM = new Random();

	public List<WordFrequency> readCNN(String namafile, List<String> myString, String warna) throws IOException {

		final FrequencyAnalyzer frequencyAnalyzer = new FrequencyAnalyzer();
		frequencyAnalyzer.setWordFrequenciesToReturn(90);
		frequencyAnalyzer.setMinWordLength(1);
		String content = new String(Files.readAllBytes(Paths
				.get("C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\src\\main\\resources\\text\\Stop Word (ALL).txt")));
		Collection<String> stopWords = new ArrayList<String>(Arrays.asList(content.split("\r\n")));
		frequencyAnalyzer.setStopWords(stopWords);

		final List<WordFrequency> wordFrequencies = frequencyAnalyzer.load(myString); // frequencyAnalyzer.load(getInputStream("text/"+FILE));
		final Dimension dimension = new Dimension(600, 600);
		final WordCloud wordCloud = new WordCloud(dimension, CollisionMode.PIXEL_PERFECT);
		wordCloud.setPadding(1);

		wordCloud.setBackground(new RectangleBackground(dimension));
		wordCloud.setBackgroundColor(new Color(255, 255, 255, 0));
		wordCloud.setKumoFont(new KumoFont("Impact", FontWeight.PLAIN));
		if (warna.equals("orange")) {
			wordCloud.setColorPalette(
					new ColorPalette(new Color(209, 64, 13), new Color(241, 83, 27), new Color(245, 130, 90),
							new Color(247, 154, 121), new Color(251, 205, 189), new Color(145, 45, 9)));
		} else {
			wordCloud.setColorPalette(new ColorPalette(new Color(0x4055F1), new Color(0x408DF1), new Color(0x40AAF1),
					new Color(0x40C5F1), new Color(0x40D3F1), new Color(0x000000)));
		}
		wordCloud.setFontScalar(new LinearFontScalar(20, 70));

		wordCloud.build(wordFrequencies);
		wordCloud.writeToFile(namafile);
		System.out.println("Wordcloud");

		return (wordFrequencies);

	}

	private static InputStream getInputStream(final String path) {
		return Thread.currentThread().getContextClassLoader().getResourceAsStream(path);
	}

}
