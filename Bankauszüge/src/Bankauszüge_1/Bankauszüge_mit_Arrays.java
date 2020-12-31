package Bankauszüge_1;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Iterator;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

public class Bankauszüge_mit_Arrays {

	public static void main(String[] args) throws IOException {
		
		String eingangxlsx = "C:\\Users\\ewegenaer.WERKSTATT-AC\\Desktop\\JavaZwischenspeicher\\bankauszugtabelle.xlsx";	// Zwischenspeicher in den csvToXLSX die Tabelle erstellt
		String ausgangxlsx = "C:\\Users\\ewegenaer.WERKSTATT-AC\\Desktop\\JavaZiel";												// holt .xlsx datei aus dem Ordner in den csvToXLSX.java sie verschiebt
		
		// XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
		
		
		switch (JOptionPane.showConfirmDialog(null, "Ist dies die richtige Datei?:\n\n" + eingangxlsx)) {
		
		case JOptionPane.YES_OPTION:
			// nix
			break;
			
		case JOptionPane.NO_OPTION:
			JFileChooser chooser1 = new JFileChooser(System.getProperty("user.home") + "\\Desktop");
			FileNameExtensionFilter filter = new FileNameExtensionFilter("XLSX", "xlsx");
			chooser1.setFileFilter(filter);
			chooser1.setAcceptAllFileFilterUsed(false);
			if(chooser1.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
				eingangxlsx = chooser1.getSelectedFile().getAbsolutePath();
			} else {
				System.out.println("Abgebrochen !!! (001)");
			}
			break;
			
		case JOptionPane.CANCEL_OPTION:
			JOptionPane.showConfirmDialog(null, "Abgebrochen! (002)", "Abbruch", JOptionPane.DEFAULT_OPTION);
			System.exit(0);
			break;
		}
		
		
		switch (JOptionPane.showConfirmDialog(null,  "Ist dies der richtige Zielordner?:\n\n" + ausgangxlsx)) {
		
		case JOptionPane.YES_OPTION:
			//nix
			break;
		
		case JOptionPane.NO_OPTION:
			JFileChooser chooser2 = new JFileChooser(System.getProperty("user.home") + "\\Desktop");
			chooser2.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
			chooser2.setAcceptAllFileFilterUsed(false);
			if(chooser2.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
				ausgangxlsx = chooser2.getSelectedFile().getAbsolutePath();
			} else {
				System.out.println("Abgebrochen !!! (003)");
			}
			break;
			
		case JOptionPane.CANCEL_OPTION:
			JOptionPane.showConfirmDialog(null, "Abgebrochen! (004)", "Abbruch", JOptionPane.DEFAULT_OPTION);
			System.exit(0);
			break;
		}
		
		
		// XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
		
		System.out.println("Eingang: " + eingangxlsx);
		System.out.println("Ausgabe: " + ausgangxlsx);
		int i = 0;																		// switch-case Indikator um Zelle zu finden; wird anfang jeder Reihe 0 gesetzt
		double betrag = 0;
		String gruppe = "";
		String kategorie = "";															// benötigt um Begünstigter in String umzuwandeln; wurde manchmal als Double erkannt
		String begünstigter = "";
		double summe = 0;
		double kategoriesonstigewert = 0;
		double gruppesonstigewert = 0;
		boolean kategoriesonstige = true;
		boolean gruppesonstige = true;
		String datumende = "";
		String datumanfang = "";
		double einnahmen = 0;
		String nix = "ÄÜÖÄÜÖÄÜÖÄÜÖ";													// Platzhalter damit das Programm nichts findet (Schutz vor false positives)
		double ausgaben = 0;
		
		
		
		// Neue Excel erstellen
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Ausgaben");
		
		// ArrayList für Namen und Werte erstellen
		
			// Kategorie
		
		ArrayList<String> namenlistekategorie = new ArrayList<>();
		namenlistekategorie.add("Bargeldauszahlung");
		namenlistekategorie.add("Bargeldauszahlung");
		namenlistekategorie.add("Kartenzahlung");
		namenlistekategorie.add("Folgelastschrift");
		
		ArrayList<String> suchwortbuchungstext = new ArrayList<>();
		suchwortbuchungstext.add("GELDAUTOMAT");											// Bargeldauszahlung
		suchwortbuchungstext.add("BARGELDAUSZAHLUNG");										// Bargeldauszahlung
		suchwortbuchungstext.add("KARTENZAHLUNG");											// Kartenzahlung
		suchwortbuchungstext.add("FOLGELASTSCHRIFT");										// Folgelastschrift
		
		ArrayList <Double> wertelistekategorie = new ArrayList<>();
		wertelistekategorie.add(0.0);														// Bargeldauszahlung
		wertelistekategorie.add(0.0);														// Bargeldauszahlung
		wertelistekategorie.add(0.0);														// Kartenzahlung
		wertelistekategorie.add(0.0);														// Folgelastschrift
		wertelistekategorie.add(0.0);
		wertelistekategorie.add(0.0);
		wertelistekategorie.add(0.0);
		wertelistekategorie.add(0.0);
		wertelistekategorie.add(0.0);
		wertelistekategorie.add(0.0);
		
			// Gruppe
		
		ArrayList<String> namenlistegruppe = new ArrayList<>();
		namenlistegruppe.add("Geldautomat");
		namenlistegruppe.add("Mama & Papa");
		namenlistegruppe.add("Lebenshilfe");
		namenlistegruppe.add("Rewe");
		namenlistegruppe.add("Edeka");
		namenlistegruppe.add("Aldi");
		namenlistegruppe.add("Kleidung");
		namenlistegruppe.add("Kleidung");
		namenlistegruppe.add("Aldi Talk");
		namenlistegruppe.add("Lieferando");
		namenlistegruppe.add("Campus Boulderhalle");
		namenlistegruppe.add("Kletterhalle Tivoli");
		namenlistegruppe.add("DAV-Duisburg");
		namenlistegruppe.add("Bars / Kneipen");
		namenlistegruppe.add("Essen gehen");
		namenlistegruppe.add("Essen gehen");
		namenlistegruppe.add("Essen gehen");
		namenlistegruppe.add("E-Scooter");
		namenlistegruppe.add("E-Scooter");
		namenlistegruppe.add("FitX");
		namenlistegruppe.add("Games");
		namenlistegruppe.add("Amazon");
		namenlistegruppe.add("Freunde");
				
		ArrayList<String> suchwortverwendungszweck = new ArrayList<>();
		suchwortverwendungszweck.add(nix);													// Geldautomat
		suchwortverwendungszweck.add(nix);													// Mama & Papa
		suchwortverwendungszweck.add("LOHN / GEHALT");										// Lebenshilfe
		suchwortverwendungszweck.add(nix);													// Rewe
		suchwortverwendungszweck.add(nix);													// Edeka
		suchwortverwendungszweck.add(nix);													// Aldi
		suchwortverwendungszweck.add(nix);													// Kleidung
		suchwortverwendungszweck.add(nix);													// Kleidung
		suchwortverwendungszweck.add("015774464877");										// Aldi Talk
		suchwortverwendungszweck.add("TAKEAWAYCOM");										// Lieferando
		suchwortverwendungszweck.add(nix);													// Campus Boulderhalle
		suchwortverwendungszweck.add(nix);													// Kletterhalle Tivoli
		suchwortverwendungszweck.add(nix);													// DAV-Duisburg
		suchwortverwendungszweck.add(nix);													// Bars / Kneipen
		suchwortverwendungszweck.add(nix);													// Essen gehen
		suchwortverwendungszweck.add(nix);													// Essen gehen
		suchwortverwendungszweck.add(nix);													// Essen gehen
		suchwortverwendungszweck.add("VOITECH");											// E-Scooter
		suchwortverwendungszweck.add("TIERMOBIL");											// E-Scooter
		suchwortverwendungszweck.add(nix);													// FitX
		suchwortverwendungszweck.add("STEAM GAMES");										// Games
		suchwortverwendungszweck.add(nix);													// Amazon
		suchwortverwendungszweck.add(nix);													// Freunde
		
		ArrayList<String> suchwortbegünstigter = new ArrayList<>();
		suchwortbegünstigter.add("GA NR00");												// Geldautomat
		suchwortbegünstigter.add("WEGENAER");												// Mama & Papa
		suchwortbegünstigter.add("Lebenshilfe");											// Lebenshilfe
		suchwortbegünstigter.add("REWE");													// Rewe
		suchwortbegünstigter.add("SCHNELLKAUF HANDELS");									// Edeka
		suchwortbegünstigter.add("ALDI");													// Aldi
		suchwortbegünstigter.add("NEW YORKER");												// Kleidung
		suchwortbegünstigter.add("GALERIA KAUFHOF");										// Kleidung
		suchwortbegünstigter.add(nix);														// Aldi Talk
		suchwortbegünstigter.add(nix);														// Lieferando
		suchwortbegünstigter.add("CAMPUS BOUL");											// Campus Boulderhalle
		suchwortbegünstigter.add("BADMINTONKLETTERHALLE TIVOL");							// Kletterhalle Tivoli
		suchwortbegünstigter.add("DAV-DUISB");												// DAV-Duisburg
		suchwortbegünstigter.add("GROTESQUE");												// Bars / Kneipen
		suchwortbegünstigter.add("LOsteria");												// Essen gehen
		suchwortbegünstigter.add("Subway");													// Essen gehen
		suchwortbegünstigter.add("1006-Aachen Aquis");										// Essen gehen
		suchwortbegünstigter.add(nix);														// E-Scooter
		suchwortbegünstigter.add(nix);														// E-Scooter
		suchwortbegünstigter.add("Fitx");													// FitX
		suchwortbegünstigter.add(nix);														// Games
		suchwortbegünstigter.add("AMAZON");													// Amazon
		suchwortbegünstigter.add("SARAH GAUTZSCH");											// Freunde
		
		ArrayList <Double> wertelistegruppe = new ArrayList<>();
		wertelistegruppe.add(0.0);															// Geldautomat
		wertelistegruppe.add(0.0);															// Mama & Papa
		wertelistegruppe.add(0.0);															// Lebenshilfe
		wertelistegruppe.add(0.0);															// Rewe
		wertelistegruppe.add(0.0);															// Edeka
		wertelistegruppe.add(0.0);															// Aldi
		wertelistegruppe.add(0.0);															// Kleidung
		wertelistegruppe.add(0.0);															// Kleidung
		wertelistegruppe.add(0.0);															// Aldi Talk
		wertelistegruppe.add(0.0);															// Lieferando
		wertelistegruppe.add(0.0);															// Campus Boulderhalle
		wertelistegruppe.add(0.0);															// Kletterhalle Tivoli
		wertelistegruppe.add(0.0);															// DAV-Duisburg
		wertelistegruppe.add(0.0);															// Bars / Kneipen
		wertelistegruppe.add(0.0);															// Essen gehen
		wertelistegruppe.add(0.0);															// Essen gehen
		wertelistegruppe.add(0.0);															// Essen gehen
		wertelistegruppe.add(0.0);															// E-Scooter
		wertelistegruppe.add(0.0);															// E-Scooter
		wertelistegruppe.add(0.0);															// FitX
		wertelistegruppe.add(0.0);															// Games
		wertelistegruppe.add(0.0);															// Amazon
		wertelistegruppe.add(0.0);															// Freunde
		wertelistegruppe.add(0.0);
		
		
		
		// Quelle Excel finden
		
		FileInputStream inputStream = new FileInputStream(new File(eingangxlsx));
		Workbook workbook2 = new XSSFWorkbook(inputStream);		
		Sheet sheet2 = workbook2.getSheetAt(0);
		
		// CellStyle
		
		CellStyle rahmenrechts = workbook.createCellStyle();
		rahmenrechts.setBorderRight(BorderStyle.MEDIUM);
		rahmenrechts.setAlignment(HorizontalAlignment.LEFT);
		
		CellStyle stylewerte = workbook.createCellStyle();
		stylewerte.setAlignment(HorizontalAlignment.RIGHT);
		
		CellStyle style1wert = workbook.createCellStyle();
		style1wert.setBorderTop(BorderStyle.THICK);
		style1wert.setAlignment(HorizontalAlignment.RIGHT);
		
		CellStyle style1name = workbook.createCellStyle();
		style1name.setBorderTop(BorderStyle.THICK);
		
		CellStyle style2 = workbook.createCellStyle();
		style2.setBorderTop(BorderStyle.DOUBLE);
		style2.setAlignment(HorizontalAlignment.RIGHT);
		
		CellStyle style2grün = workbook.createCellStyle();
		style2grün.setBorderTop(BorderStyle.DOUBLE);
		style2grün.setFillForegroundColor(IndexedColors.BRIGHT_GREEN1.getIndex());
		style2grün.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style2grün.setAlignment(HorizontalAlignment.RIGHT);
		
		CellStyle style2rot = workbook.createCellStyle();
		style2rot.setBorderTop(BorderStyle.DOUBLE);
		style2rot.setFillForegroundColor(IndexedColors.RED.getIndex());
		style2rot.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style2rot.setAlignment(HorizontalAlignment.RIGHT);
		
		CellStyle rahmenunten = workbook.createCellStyle();
		rahmenunten.setBorderBottom(BorderStyle.THICK);
		
		CellStyle linksbündigsumme = workbook.createCellStyle();
		linksbündigsumme.setBorderTop(BorderStyle.DOUBLE);
		linksbündigsumme.setAlignment(HorizontalAlignment.LEFT);
		
		CellStyle linksbündig = workbook.createCellStyle();
		linksbündig.setBorderTop(BorderStyle.THICK);
		linksbündig.setAlignment(HorizontalAlignment.LEFT);
		
		// Zählungsiteratoren
		
		Iterator<Row> iterator = sheet2.iterator();
		Iterator<Row> iterator2 = sheet2.iterator();
		
		
		int reihen = 0;
		int reihengesamt = 0;
		while (iterator.hasNext()) {
			iterator.next();
			reihen++;
		}
		
		reihengesamt = reihen;
		System.out.println("Gesamtreihen: " + reihengesamt);
		
		
		
		
		while (reihen > 0) {											// Reihendurchlauf Beginn
			
			Row nextRow = iterator2.next();
			Iterator<Cell> cellIterator2 = nextRow.cellIterator();
			System.out.println("------------------------------------------------------------------------------------------------");
			System.out.println("Reihe Nummer: " + (reihengesamt - reihen + 1));
			i = 0;																// Zähler in der Reihe um Zelle zu finden: [3]"Buchungstext", [4]"Verwendungszweck", [11]"Begünstigter/Zahlunspflichtiger", [14]"Betrag"
			betrag = 0;
			gruppe = "";														// Essensbestellung, Rewe Kartenzahlung, etc...
			kategorie = "";														// Kartenzahlung, Bargeldauszahlung, Sonstige Zahlung
			kategoriesonstige = true;
			gruppesonstige = true;
			
			for (int i2 = 0; i2 < 17; i2++) {									// Zellendurchlauf in Reihe Beginn| while (cellIterator2.hasNext()) {
				
				Cell cell = cellIterator2.next();
				
				switch (i) {													// Test in Welche Spalte wir uns befinden
					
					case (1):
						if ((reihengesamt - reihen + 1) == 2) {
							datumende = cell.getStringCellValue();
						}
						if (reihen == 1) {
							datumanfang = cell.getStringCellValue();
						}
				
					case (3):																	// Zelle 3 in der Reihe ist "Buchungstext"
						for (int i001 = 0; i001 < suchwortbuchungstext.size(); i001++) {
							if (cell.getStringCellValue().contains(suchwortbuchungstext.get(i001))) {
								kategorie = namenlistekategorie.get(i001);
								kategoriesonstige = false;
							}
						}
						break;
					
					case (4):																	// Zelle 4 in der Reihe ist "Verwendungszweck"
						for (int i002 = 0; i002 < suchwortverwendungszweck.size(); i002++) {
							if (cell.getStringCellValue().contains(suchwortverwendungszweck.get(i002))) {
								gruppe = namenlistegruppe.get(i002);
								gruppesonstige = false;
							}
						}
						break;
					
						
					case (11):																	// Zelle 11 der Reihe ist "Begünstigter/Zahlungspflichtiger"
						
						switch (cell.getCellType()) {
							case NUMERIC:
								begünstigter = Double.toString(cell.getNumericCellValue());
								break;
							case STRING:
								begünstigter = cell.getStringCellValue();
								break;
							default:
						}
					
						for (int i003 = 0; i003 < suchwortbegünstigter.size(); i003++) {
							if (begünstigter.contains(suchwortbegünstigter.get(i003))) {
								gruppe = namenlistegruppe.get(i003);
								gruppesonstige = false;
							}
						}
						break;
						
					
					case (14):																	// Zelle 14 der Reihe ist "Betrag"
						
						if (cell.getCellType() == CellType.STRING) {							// Damit in der ersten Reihe mit "Betrag" keine Fehler auftreten
							String text = cell.getStringCellValue();							// Falls eine Zahl als String gelesen wird, wird sie hier in Double umgewandelt
							if (!text.equals("Betrag")) {
								text = text.replace(".", "");
								text = text.replace(",", ".");
								betrag = Double.valueOf(text);
							}
						} else {
						betrag = cell.getNumericCellValue();
						}
						summe = Math.round((summe + betrag) * 100d) / 100d;
						if (betrag < 0) {
							ausgaben = Math.round((ausgaben + betrag) *100d) /100d;
						}
						break;
					default:
				}
				
				i = i + 1;
				
			}																	// Zellendurchlauf in Reihe Ende
			
			
			// Am Ende von jeder Reihe den Betrag auf in das wertearray addieren
			
			if (!kategoriesonstige) {
				for (int i004 = 0; i004 < namenlistekategorie.size(); i004++) {
					if (kategorie.equals(namenlistekategorie.get(i004))) {
						wertelistekategorie.set(i004, (wertelistekategorie.get(i004) + betrag));
					}
				}
			} else {
				if (betrag < 0) {
					kategorie = "Sonstige Kategorie";
					kategoriesonstigewert = kategoriesonstigewert + betrag;
				} else {
					kategorie = "Betrag ist positiv!";
					einnahmen = einnahmen + betrag;
					
				}
			}
			
			if (!gruppesonstige) {
				for (int i005 = 0; i005 < namenlistegruppe.size(); i005++) {
					if (gruppe.equals(namenlistegruppe.get(i005))) {
						wertelistegruppe.set(i005,  (wertelistegruppe.get(i005) + betrag));
					}
				}
			} else {
				if (betrag < 0) {
					gruppe = "Sonstige Gruppe";
					gruppesonstigewert = gruppesonstigewert + betrag;
				} else {
					gruppe = "Betrag ist positiv!";
				}
			}
			
			// Am Ende von jeder Reihe werden alle Zahlen gerundet (Einige runde Additionen gaben viele Nachkommastellen aus)

			for (int n3 = 0; n3 < wertelistekategorie.size(); n3++) {
				wertelistekategorie.set(n3, Math.round(wertelistekategorie.get(n3) *100d) / 100d);
			}
			
			for (int n4 = 0; n4 < wertelistegruppe.size(); n4++) {
				wertelistegruppe.set(n4, Math.round(wertelistegruppe.get(n4) *100d) / 100d);
			}
			
			kategoriesonstigewert = Math.round(kategoriesonstigewert *100d) / 100d;
			gruppesonstigewert = Math.round(gruppesonstigewert *100d) / 100d;
			ausgaben = Math.round(ausgaben *100d) / 100d;
			
			reihen = reihen - 1;
						
			System.out.println("Kategorie: \t\t" + kategorie);
			System.out.println("Gruppe: \t\t" + gruppe);
			System.out.println(reihen);
			System.out.println("Kategorie\t" + namenlistekategorie + "\tSonstige:");
			System.out.println("Kategorie\t" + wertelistekategorie + "\t" + kategoriesonstigewert);
			System.out.println("Gruppe\t\t" + namenlistegruppe + "\tSonstige");
			System.out.println("Gruppe\t\t" + wertelistegruppe + "\tSonstige: " + gruppesonstigewert);
			System.out.println("Betrag:\t " + betrag);
			System.out.println("Summe: " + summe);
		}																		// Reihendurchlauf Ende
		
		
		// Datum und Tage ausrechnen und schreiben
		
		System.out.println("Enddatum: " + datumende);
		System.out.println("Anfangsdatum: " + datumanfang);
		Row reihe5 = sheet.createRow(0);
		Cell zelle9 = reihe5.createCell(0);
		zelle9.setCellValue("Anfangsdatum");
		Cell zelle10 = reihe5.createCell(1);
		zelle10.setCellValue(datumanfang);
		zelle10.setCellStyle(stylewerte);
		
		Row reihe6 = sheet.createRow(1);
		Cell zelle11 = reihe6.createCell(0);
		zelle11.setCellValue("Enddatum");
		Cell zelle12 = reihe6.createCell(1);
		zelle12.setCellValue(datumende);
		zelle12.setCellStyle(stylewerte);
		
		datumanfang = datumanfang.replace(".", "");
		datumende = datumende.replace(".", "");
		
		DateTimeFormatter dtf2 = DateTimeFormatter.ofPattern("ddMMyy");
		DateTimeFormatter dtf3 = DateTimeFormatter.ofPattern("ddMMyyyy");
		LocalDate datumanfangdtf = null;
		LocalDate datumendedtf = null;
		
		if (datumanfang.length() == 6) {
			datumanfangdtf = LocalDate.parse(datumanfang, dtf2);
			datumendedtf = LocalDate.parse(datumende, dtf2);
		} else {
			datumanfangdtf = LocalDate.parse(datumanfang, dtf3);
			datumendedtf = LocalDate.parse(datumende, dtf3);
		}
		double tage = Duration.between(datumanfangdtf.atStartOfDay(), datumendedtf.atStartOfDay()).toDays() + 1;			// + 1 weil im bankauszug Anfangs- und Enddatum inklusive sind
		
		Cell zelle13 = reihe6.createCell(3);
		zelle13.setCellValue("Tage:");
		Cell zelle14 = reihe6.createCell(4);
		zelle14.setCellValue(tage);
		
		// Werte in Tabelle schreiben
		
		// Kategorie
		
		int n1x = 5; 																										// erste Reihe ist die 3, damit Platz für Anfangs un Enddatum ist
		int n1xpos = n1x;
		int n2x = 0;
		
		for (int n1 = 0; n1 < namenlistekategorie.size(); n1++) {
			if (n1 != 0) {
				if (namenlistekategorie.get(n1).equals(namenlistekategorie.get(n1-1))) {									// Nicht ändern, mit !namenlistekategorie.get(n1).equals(namenlistekategorie.get(n1-1)) als if Bedingung funtioniert nix mehr
					// Nichts
				} else {
					if (wertelistekategorie.get(n1) < 0) {
						if (sheet.getRow(n1x) == null) {
							Row reihe1 = sheet.createRow(n1x);
							reihe1.setHeight((short) 400);
							Cell zelle1 = reihe1.createCell(0);
							zelle1.setCellValue(namenlistekategorie.get(n1));
							Cell zelle2 = reihe1.createCell(1);
							zelle2.setCellValue(wertelistekategorie.get(n1));
							zelle2.setCellStyle(stylewerte);
							Cell zelle3 = reihe1.createCell(2);
							zelle3.setCellValue("EUR");
							n1x = n1x + 1;
						} else {
							Cell zelle1 = sheet.getRow(n1x).createCell(0);
							zelle1.setCellValue(namenlistekategorie.get(n1));
							Cell zelle2 = sheet.getRow(n1x).createCell(1);
							zelle2.setCellValue(wertelistekategorie.get(n1));
							zelle2.setCellStyle(stylewerte);
							Cell zelle3 = sheet.getRow(n1x).createCell(2);
							zelle3.setCellValue("EUR");;
							n1x = n1x + 1;
						}
					} else if (wertelistekategorie.get(n1) > 0) {
						if (sheet.getRow(n1xpos) == null) {
							Row reihe2 = sheet.createRow(n1xpos);
							reihe2.setHeight((short) 400);
							Cell zelle3 = reihe2.createCell(3);
							zelle3.setCellValue(namenlistekategorie.get(n1));
							Cell zelle4 = reihe2.createCell(4);
							zelle4.setCellValue(wertelistekategorie.get(n1));
							zelle4.setCellStyle(stylewerte);
							Cell zelle5 = reihe2.createCell(2);
							zelle5.setCellValue("EUR");
							n1xpos = n1xpos + 1;
						} else {
							Cell zelle3 = sheet.getRow(n1xpos).createCell(3);
							zelle3.setCellValue(namenlistekategorie.get(n1));
							Cell zelle4 = sheet.getRow(n1xpos).createCell(4);
							zelle4.setCellValue(wertelistekategorie.get(n1));
							zelle4.setCellStyle(stylewerte);
							Cell zelle5 = sheet.getRow(n1xpos).createCell(5);
							zelle5.setCellValue("EUR");
							n1xpos = n1xpos + 1;
						}
					}
				}
			} else {
				if (wertelistekategorie.get(n1) < 0) {
					Row reihe1 = sheet.createRow(n1x);
					reihe1.setHeight((short) 400);
					Cell zelle1 = reihe1.createCell(0);
					zelle1.setCellValue(namenlistekategorie.get(n1));
					Cell zelle2 = reihe1.createCell(1);
					zelle2.setCellValue(wertelistekategorie.get(n1));
					zelle2.setCellStyle(stylewerte);
					Cell zelle3 = reihe1.createCell(2);
					zelle3.setCellValue("EUR");
					n1x = n1x + 1;
				} else if (wertelistekategorie.get(n1) > 0) {
					if (sheet.getRow(n1xpos) == null) {
						Row reihe2 = sheet.createRow(n1xpos);
						reihe2.setHeight((short) 400);
						Cell zelle3 = reihe2.createCell(3);
						zelle3.setCellValue(namenlistekategorie.get(n1));
						Cell zelle4 = reihe2.createCell(4);
						zelle4.setCellValue(wertelistekategorie.get(n1));
						zelle4.setCellStyle(stylewerte);
						Cell zelle5 = reihe2.createCell(5);
						zelle5.setCellValue("EUR");
						n1xpos = n1xpos + 1;
					} else {
						Cell zelle3 = sheet.getRow(n1xpos).createCell(3);
						zelle3.setCellValue(namenlistekategorie.get(n1));
						Cell zelle4 = sheet.getRow(n1xpos).createCell(4);
						zelle4.setCellValue(wertelistekategorie.get(n1));
						zelle4.setCellStyle(stylewerte);
						Cell zelle5 = sheet.getRow(n1xpos).createCell(2);
						zelle5.setCellValue("EUR");
						n1xpos = n1xpos + 1;
					}
				}
				
			}
			
		}
		
		// Sonstige Kategorie
		
		if (kategoriesonstigewert < 0) {
			Row reihe3 = sheet.createRow(n1x);
			reihe3.setHeight((short) 400);
			Cell zelle5 = reihe3.createCell(0);
			zelle5.setCellValue("Sonstiges");
			Cell zelle6 = reihe3.createCell(1);
			zelle6.setCellValue(kategoriesonstigewert);
			zelle6.setCellStyle(stylewerte);
			Cell zelle7 = reihe3.createCell(2);
			zelle7.setCellValue("EUR");			
		} else {
			Row reihe3 = sheet.createRow(n1x);
			reihe3.setHeight((short) 400);
			Cell zelle5 = reihe3.createCell(3);
			zelle5.setCellValue("Sonstiges");
			Cell zelle6 = reihe3.createCell(4);
			zelle6.setCellValue(kategoriesonstigewert);
			zelle6.setCellStyle(stylewerte);
			Cell zelle7 = reihe3.createCell(5);
			zelle7.setCellValue("EUR");
			
		}
		
		if (n1x < n1xpos) {
			n1x = n1xpos;
		}
		n2x = n1x + 2;	
		int n2xpos = n2x;
		
		
		// Gruppe
		
		for (int n2 = 0; n2 < namenlistegruppe.size(); n2++) {
			if (n2 != 0) {
				if (namenlistegruppe.get(n2).equals(namenlistegruppe.get(n2 - 1))) {
					// Nichts
				} else {
					if (wertelistegruppe.get(n2) < 0) {
						if (sheet.getRow(n2x) == null) {
							Row reihe2 = sheet.createRow(n2x);
							reihe2.setHeight((short) 400);
							Cell zelle3 = reihe2.createCell(0);
							zelle3.setCellValue(namenlistegruppe.get(n2));
							Cell zelle4 = reihe2.createCell(1);
							zelle4.setCellValue(wertelistegruppe.get(n2));
							zelle4.setCellStyle(stylewerte);
							Cell zelle5 = reihe2.createCell(2);
							zelle5.setCellValue("EUR");
							n2x = n2x + 1;
						} else {
							Cell zelle3 = sheet.getRow(n2x).createCell(0);
							zelle3.setCellValue(namenlistegruppe.get(n2));
							Cell zelle4 = sheet.getRow(n2x).createCell(1);
							zelle4.setCellValue(wertelistegruppe.get(n2));
							zelle4.setCellStyle(stylewerte);
							Cell zelle5 = sheet.getRow(n2x).createCell(2);
							zelle5.setCellValue("EUR");
							n2x = n2x + 1;
						}
					} else if (wertelistegruppe.get(n2) > 0) {
						if (sheet.getRow(n2xpos) == null) {
							Row reihe2 = sheet.createRow(n2xpos);
							reihe2.setHeight((short) 400);
							Cell zelle3 = reihe2.createCell(3);
							zelle3.setCellValue(namenlistegruppe.get(n2));
							Cell zelle4 = reihe2.createCell(4);
							zelle4.setCellValue(wertelistegruppe.get(n2));
							zelle4.setCellStyle(stylewerte);
							Cell zelle5 = reihe2.createCell(5);
							zelle5.setCellValue("EUR");
							n2xpos = n2xpos + 1;
						} else {
							Cell zelle3 = sheet.getRow(n2xpos).createCell(3);
							zelle3.setCellValue(namenlistegruppe.get(n2));
							Cell zelle4 = sheet.getRow(n2xpos).createCell(4);
							zelle4.setCellValue(wertelistegruppe.get(n2));
							zelle4.setCellStyle(stylewerte);
							Cell zelle5 = sheet.getRow(n2xpos).createCell(5);
							zelle5.setCellValue("EUR");
							n2xpos = n2xpos + 1;
						}
					}
				}
			} else {
				if (wertelistegruppe.get(n2) < 0) {
					Row reihe2 = sheet.createRow(n2x);
					reihe2.setHeight((short) 400);
					Cell zelle3 = reihe2.createCell(0);
					zelle3.setCellValue(namenlistegruppe.get(n2));
					Cell zelle4 = reihe2.createCell(1);
					zelle4.setCellValue(wertelistegruppe.get(n2));
					zelle4.setCellStyle(stylewerte);
					Cell zelle5 = reihe2.createCell(2);
					zelle5.setCellValue("EUR");
					n2x = n2x + 1;
				} else if (wertelistegruppe.get(n2) > 0) {
					if (sheet.getRow(n2xpos) == null) {
						Row reihe2 = sheet.createRow(n2xpos);
						reihe2.setHeight((short) 400);
						Cell zelle3 = reihe2.createCell(3);
						zelle3.setCellValue(namenlistegruppe.get(n2));
						Cell zelle4 = reihe2.createCell(4);
						zelle4.setCellValue(wertelistegruppe.get(n2));
						zelle4.setCellStyle(stylewerte);
						Cell zelle5 = reihe2.createCell(5);
						zelle5.setCellValue("EUR");
						n2xpos = n2xpos + 1;
					} else {
						Cell zelle3 = sheet.getRow(n2xpos).createCell(3);
						zelle3.setCellValue(namenlistegruppe.get(n2));
						Cell zelle4 = sheet.getRow(n2xpos).createCell(4);
						zelle4.setCellValue(wertelistegruppe.get(n2));
						zelle4.setCellStyle(stylewerte);
						Cell zelle5 = sheet.getRow(n2xpos).createCell(5);
						zelle5.setCellValue("EUR");
						n2xpos = n2xpos + 1;
					}
				}
			}
		}
		
		// Sonstige Gruppe
		
		if (n2x < n2xpos) {
			n2x = n2xpos;
		}
		
		
		if (gruppesonstigewert < 0) {
			Row reihe4 = sheet.createRow(n2x);
			reihe4.setHeight((short) 400);
			Cell zelle7 = reihe4.createCell(0);
			zelle7.setCellValue("Sonstiges");
			Cell zelle8 = reihe4.createCell(1);
			zelle8.setCellValue(gruppesonstigewert);
			zelle8.setCellStyle(stylewerte);
			Cell zelle15 = reihe4.createCell(2);
			zelle15.setCellValue("EUR");
		} else {
			Row reihe4 = sheet.createRow(n2x);
			reihe4.setHeight((short) 400);
			Cell zelle7 = reihe4.createCell(3);
			zelle7.setCellValue("Sonstiges");
			Cell zelle8 = reihe4.createCell(4);
			zelle8.setCellValue(gruppesonstigewert);
			zelle8.setCellStyle(stylewerte);
			Cell zelle15 = reihe4.createCell(2);
			zelle15.setCellValue("EUR");
		}
		
		
		
		// Rahmen über Gruppe
		Row reihe10 = sheet.createRow(n1x + 1);
		reihe10.setHeight((short) 250);
		for (int i6 = 0; i6 < 6; i6++) {
			Cell zelle22 = reihe10.createCell(i6);
			zelle22.setCellStyle(rahmenunten);
		}
		
		// Rahmen rechts von Gruppe
		for (int i7 = (n1x + 2); i7 < (n2x + 1); i7++) {
			if (sheet.getRow(i7).getCell(5) == null) {
				Cell zelle22 = sheet.getRow(i7).createCell(5);
				zelle22.setCellStyle(rahmenrechts);
			} else {
				sheet.getRow(i7).getCell(5).setCellStyle(rahmenrechts);
			}
		}
		
		// Rahmen letzte Zeile von Gruppe
		Row reihe11 = sheet.createRow(n2x + 1);
		reihe11.setHeight((short) 250);
		Cell zelle23 = reihe11.createCell(5);
		zelle23.setCellStyle(rahmenrechts);
		
		Row reihe9 = sheet.createRow(n2x + 2);
		reihe9.setHeight((short) 400);
		Cell zelle19 = reihe9.createCell(0);
		zelle19.setCellStyle(style1name);
		zelle19.setCellValue("Ausgaben:");
		Cell zelle20 = reihe9.createCell(1);
		zelle20.setCellStyle(style1wert);
		zelle20.setCellValue(ausgaben);
		Cell zelle22 = reihe9.createCell(2);
		zelle22.setCellValue("EUR");
		zelle22.setCellStyle(style1wert);
		
		Cell zelle21 = reihe9.createCell(2);
		zelle21.setCellValue("EUR");
		zelle21.setCellStyle(style1name);
		
		
		Cell zelle15 = reihe9.createCell(3);
		zelle15.setCellStyle(style1name);
		zelle15.setCellValue("Einnahmen:");
		Cell zelle16 = reihe9.createCell(4);
		zelle16.setCellValue(einnahmen);
		zelle16.setCellStyle(style1wert);
		Cell zelle24 = reihe9.createCell(5);
		zelle24.setCellValue("EUR");
		zelle24.setCellStyle(linksbündig);
		
		Row reihe8 = sheet.createRow(n2x + 4);
		reihe8.setHeight((short) 400);
		Cell zelle17 = reihe8.createCell(3);
		zelle17.setCellStyle(style2);
		zelle17.setCellValue("Summe:");
		Cell zelle18 = reihe8.createCell(4);
		zelle18.setCellStyle(stylewerte);
		zelle18.setCellValue(summe);
		Cell zelle25 = reihe8.createCell(5);
		zelle25.setCellValue("EUR");
		zelle25.setCellStyle(linksbündigsumme);
		
		if (summe >= 0) {
			zelle18.setCellStyle(style2grün);
		} else {
			zelle18.setCellStyle(style2rot);
		}
		
		sheet.autoSizeColumn(0);
		sheet.autoSizeColumn(1);
		sheet.autoSizeColumn(3);
		sheet.autoSizeColumn(4);
		sheet.setColumnWidth(2, 3000);
		sheet.setColumnWidth(5, 3000);
		
		
		// Datensatz für PieChart in Excel schreiben (in weiß in Zelle A3)
		
		String datarangeausgabennamen = "A11:A" + (n2x + 1);
		String datarangeausgabenwerte = "B11:B" + (n2x + 1);
		String datarangeeinnahmennamen = "D11:D" + (n2x + 1);
		String datarangeeinnahmenwerte = "E11:E" + (n2x + 1);
		System.out.println("datarange ist: " + datarangeausgabennamen);
		
				
		CellStyle weiß = workbook.createCellStyle();
		Font weißfont = workbook.createFont();
		weißfont.setColor(IndexedColors.WHITE.getIndex());
		weiß.setFont(weißfont);
		sheet.createRow(2).createCell(0).setCellValue(datarangeausgabennamen);
		sheet.getRow(2).getCell(0).setCellStyle(weiß);
		sheet.getRow(2).createCell(1).setCellValue(datarangeausgabenwerte);
		sheet.getRow(2).getCell(1).setCellStyle(weiß);
		sheet.getRow(2).createCell(2).setCellValue(datarangeeinnahmennamen);
		sheet.getRow(2).getCell(2).setCellStyle(weiß);
		sheet.getRow(2).createCell(3).setCellValue(datarangeeinnahmenwerte);
		sheet.getRow(2).getCell(3).setCellStyle(weiß);
		
		
		// Einmaligen Namen für neue Tabelle erschaffen
		
		DateTimeFormatter dtf4 = DateTimeFormatter.ofPattern("dd.LLL");
		String datumname = dtf4.format(datumanfangdtf) + " bis " + dtf4.format(datumendedtf);
		
		DateTimeFormatter dtf1 = DateTimeFormatter.ofPattern("HHmmssSS");
		LocalDateTime now = LocalDateTime.now();
		String nummer = dtf1.format(now);
		String dateiname = "\\Auswertung " + datumname + " " + nummer + ".xlsx";
		String ziel = ausgangxlsx + dateiname;
		
		// Tabelle speichern
		
		try {
			FileOutputStream output = new FileOutputStream(ziel);
			workbook.write(output);
			output.close();
		} catch(Exception e) {
			e.printStackTrace();
		}
		
		try {
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
				
		workbook2.close();
		workbook.close();
		
		JOptionPane.showConfirmDialog(null, "Tabelle wurde analysiert !\nDie Ergebnisse:\n\n" + ziel, "Erfolgreich", JOptionPane.DEFAULT_OPTION);
		Desktop desktop = Desktop.getDesktop();
		try {
			File datei = new File(ausgangxlsx);
			desktop.open(datei);
		} catch (IllegalArgumentException iae) {
			System.out.println("Fehler (005)");
		}
		
	}

}
