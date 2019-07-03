package com.ebdesk.report.slide;

import java.awt.Color;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Scanner;
import java.util.Set;

import javax.imageio.ImageIO;

import org.apache.poi.hslf.usermodel.HSLFLine;
import org.apache.poi.hslf.usermodel.HSLFTable;
import org.apache.poi.hslf.usermodel.HSLFTableCell;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.sl.usermodel.VerticalAlignment;

import javafx.embed.swing.SwingFXUtils;
import javafx.scene.SnapshotParameters;
import javafx.scene.image.WritableImage;

public class percobaan {
	private static final char DEFAULT_SEPARATOR = ';';
	private static final char DEFAULT_QUOTE = '"';

	public static void main(String[] args) throws FileNotFoundException {
		// ====================CSV==================//

		// ====================CSV==================//
		String csvFile = "C:\\Users\\eBdesk\\Desktop\\JAVA\\sna-report\\output\\top_account_by_reach_Agus Yudhoyono.csv";

		BufferedReader br = null;
		String line = "";
		String cvsSplitBy = ",";

		int a = 4;
		int b = 4;
		int xx = 0;
		int xc = 0;

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

}
