package Bankauszüge_1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Bankauszüge_sortieren {

	public static void main(String[] args) throws IOException {
		
		/*
		 * Gelöste Probleme:
		 * - null Zellen: Beide while Schleifen machen nun eine fest Zahl an Durchläufen
		 * - Liest Betrag Spalte als String: Tut es immernoch, String wird aber in Double umgeformt
		 * - Beträge sind mit Kommas geschrieben: Kommas werden nach dem Lesen durch Punkt ersetzt
		 * - Wenn Beträge ab 1000€ ein Tausenderpunkt haben, wird das Programm diesen als Kommastelle lesen: Erstzt Punkte durch Kommas und danach Kommas durch Punkte
		 * 
		 * Probleme:
		 * 
		 * Braucht:
		 * .xlsx Format
		 * 
		 */
		
		
		String zwischenspeicher = "C:\\Users\\ewegenaer.WERKSTATT-AC\\Desktop\\JavaZwischenspeicher\\bankauszugtabelle.xlsx";
		String ausgang = "C:\\Users\\ewegenaer.WERKSTATT-AC\\Desktop\\JavaZiel\\";					// holt .xlsx datei aus dem Ordner in den csvToXLSX.java sie verschiebt
		System.out.println("Eingang: " + zwischenspeicher);
		System.out.println("Ausgabe: " + ausgang);
		int i = 0;
		double betrag = 0;
		String gruppe = "";
		String kategorie = "";
		String test002 = "";
		double summe = 0;
		
		//Kategoriesummen:
		
		double bargeldauszahlung = 0;
		double kartenzahlung = 0;
		double sonstiges = 0;
		
		//Gruppensummen
		
		double rewe = 0;
		double geldautomat = 0;
		double edeka = 0;
		double sonstigekartenzahlung = 0;
		
		
		// Neue Excel erstellen
		
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Ausgaben");
		
		// Layout von neuer Excel
		
		Row reihe2 = sheet.createRow(1);
		reihe2.setHeight((short) 500);
		Row reihe3 = sheet.createRow(2);
		reihe3.setHeight((short) 500);
		Row reihe4 = sheet.createRow(3);
		reihe4.setHeight((short) 500);
		Row reihe6 = sheet.createRow(5);
		reihe4.setHeight((short) 500);
		
		Cell zelleB2 = reihe2.createCell(1);
		Cell zelleH2= reihe2.createCell(7);
		Cell zelleB3 = reihe3.createCell(1);
		Cell zelleC3 = reihe3.createCell(2);
		Cell zelleD3 = reihe3.createCell(3);
		Cell zelleH3 = reihe3.createCell(7);
		Cell zelleI3 = reihe3.createCell(8);
		Cell zelleJ3 = reihe3.createCell(9);
		Cell zelleK3 = reihe3.createCell(10);
		
		Cell zelleB4 = reihe4.createCell(1);
		Cell zelleC4 = reihe4.createCell(2);
		Cell zelleD4 = reihe4.createCell(3);
		Cell zelleH4 = reihe4.createCell(7);
		Cell zelleI4 = reihe4.createCell(8);
		Cell zelleJ4 = reihe4.createCell(9);
		Cell zelleK4 = reihe4.createCell(10);
		
		Cell zelleB6 = reihe6.createCell(1);
		Cell zelleC6 = reihe6.createCell(2);
		
		zelleB2.setCellValue("Kategorie:");
		zelleH2.setCellValue("Gruppe");
		zelleB3.setCellValue("Bargeldauszahlung");
		zelleC3.setCellValue("Kartenzahlung");
		zelleD3.setCellValue("Sonstiges");
		zelleH3.setCellValue("Rewe");
		zelleI3.setCellValue("Geldautomat");
		zelleJ3.setCellValue("Edeka");
		zelleK3.setCellValue("Sonstige Kartenzahlung");
		zelleB6.setCellValue("Netto:");
		
		
		// Quelle Excel finden
		
		FileInputStream inputStream = new FileInputStream(new File(zwischenspeicher));
		Workbook workbook2 = new XSSFWorkbook(inputStream);		
		Sheet sheet2 = workbook2.getSheetAt(0);
		
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
		System.out.println("reihen: " + reihengesamt);
		
		
		
		
		while (reihen > 0) {											// Reihendurchlauf Beginn
			
			Row nextRow = iterator2.next();
			Iterator<Cell> cellIterator2 = nextRow.cellIterator();
			System.out.println("Reihe Nummer: " + (reihengesamt - reihen + 1));
			i = 0;																// Zähler in der Reihe um Zelle zu finden: [3]"Buchungstext", [4]"Verwendungszweck", [11]"Begünstigter/Zahlunspflichtiger", [14]"Betrag"
			betrag = 0;
			gruppe = "";														// Essensbestellung, Rewe Kartenzahlung, etc...
			kategorie = "";														// Kartenzahlung, Bargeldauszahlung, Sonstige Zahlung
			
			for (int i2 = 0; i2 < 17; i2++) {									// Zellendurchlauf in Reihe Beginn| while (cellIterator2.hasNext()) {
				
				Cell cell = cellIterator2.next();
				switch (i) {													// Test in Welche Spalte wir uns befinden
					
					case (3):														// Zelle 3 in der Reihe ist "Buchungstext"
						switch (cell.getStringCellValue()) {
							case ("GELDAUTOMAT"): case ("BARGELDAUSZAHLUNG"):
								kategorie = "Bargeldauszahlung";
								break;
							case ("KARTENZAHLUNG"):
								kategorie = "Kartenzahlung";
								break;
							default:
								kategorie = "Sonstige Zahlung";
						}
						break;
					
					case (4):														// Zelle 4 in der Reihe ist "Verwendungszweck"
						if (cell.getStringCellValue().contains("TAKEAWAYCOM")) {
							gruppe = "Bestelltes Essen";
						}
						break;
						
						
						
					case (11):
						
						switch (cell.getCellType()) {
							case NUMERIC:
								test002 = Double.toString(cell.getNumericCellValue());
								break;
							case STRING:
								test002 = cell.getStringCellValue();
								break;
							default:
						}
					
						// System.out.print(test002 +"\n");
					
						if (test002.contains("REWE")) {
							gruppe = "Rewe";
						} else if (test002.contains("GA NR00")) {
							gruppe = "Geldautomat";
						} else if (test002.contains("SCHNELLKAUF HANDELS")) {
							gruppe = "Edeka";
						} else {
							gruppe = "Sonstige Kartenzahlung";
						}
						break;
						
					
					case (14):
						
						if (cell.getCellType() == CellType.STRING) {								// Damit in der ersten Reihe mit "Betrag" keine Fehler auftreten
							//System.out.println("Fehler: Betragszelle als " + cell.getCellType() + " erkannt!!!: " + cell.getStringCellValue());
							String text = cell.getStringCellValue();
							if (!text.equals("Betrag")) {
								System.out.println("1.     " + text);
								text = text.replace(".", "");
								System.out.println("2.     " + text);
								text = text.replace(",", ".");
								System.out.println("3.     " + text);
								betrag = Double.valueOf(text);
							}
						} else {
						betrag = cell.getNumericCellValue();
						}
						summe = summe + betrag;
						break;
					default:
				}
				
				/*
				if (cell.getCellType() == CellType.NUMERIC) {
					System.out.print(cell.getNumericCellValue() + "/");
				} else {
					System.out.print(cell.getStringCellValue() + "/");
				}
				*/
				
				System.out.print("Kategorie: " + kategorie + "|| Gruppe: " + gruppe + "|| Betrag: " + betrag + "|| Summe: " + summe + "\n");
				
				i = i + 1;
				
			}																	// Zellendurchlauf in Reihe Ende
			
			
			
			
			
			
			switch (kategorie) {
				case ("Bargeldauszahlung"):
					bargeldauszahlung = bargeldauszahlung + betrag;
					break;
				case ("Kartenzahlung"):
					kartenzahlung = kartenzahlung + betrag;
					break;
				case ("Sonstige Zahlung"):
					if (betrag < 0) {
						sonstiges = sonstiges + betrag;
					}
					break;
				default:
			}
			
			switch (gruppe) {
				case ("Rewe"):
					rewe = rewe + betrag;
					break;
				case ("Geldautomat"):
					geldautomat = geldautomat + betrag;
					break;
				case ("Edeka"):
					edeka = edeka + betrag;
					break;
				case ("Sonstige Kartenzahlung"):
					if (betrag < 0) {
						sonstigekartenzahlung = sonstigekartenzahlung + betrag;
					}
					break;
				default:
			}
			
			reihen = reihen - 1;
		}																		// Reihendurchlauf Ende
		
		
		bargeldauszahlung = Math.round(bargeldauszahlung * 100d) / 100d;
		kartenzahlung = Math.round(kartenzahlung * 100d) / 100d;
		sonstiges = Math.round(sonstiges * 100d) / 100d;
		rewe = Math.round(rewe * 100d) / 100d;
		geldautomat = Math.round(geldautomat * 100d) / 100d;
		edeka = Math.round(edeka * 100d) / 100d;
		sonstigekartenzahlung = Math.round(sonstigekartenzahlung * 100d) / 100d;
		summe = Math.round(summe * 100d) / 100d;
		
		
		zelleB4.setCellValue(bargeldauszahlung + "€");
		zelleC4.setCellValue(kartenzahlung + "€");
		zelleD4.setCellValue(sonstiges + "€");
		zelleH4.setCellValue(rewe + "€");
		zelleI4.setCellValue(geldautomat + "€");
		zelleJ4.setCellValue(edeka + "€");
		zelleK4.setCellValue(sonstigekartenzahlung + "€");
		zelleC6.setCellValue(summe + "€");
		

		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy dd.LLL HH-mm-ss");
		LocalDateTime now = LocalDateTime.now();
		String nummer = dtf.format(now);
		String dateiname = "Ausgaben " + nummer + ".xlsx";
		
		try {
			FileOutputStream output = new FileOutputStream(dateiname);
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
		
		
		
		String quellpfad = "C:\\Users\\ewegenaer.WERKSTATT-AC\\OneDrive\\Java Projekte\\AusgabeTests\\";
		String quelle = quellpfad + dateiname;
		
		String zielpfad = ausgang;
		String ziel = zielpfad + dateiname;
		
	      moveFile(quelle, ziel);
		
		
		
		workbook2.close();
		workbook.close();
		

	}
	
	 private static void moveFile(String src, String dest ) {
	      Path result = null;
	      try {
	         result = Files.move(Paths.get(src), Paths.get(dest));
	      } catch (IOException e) {
	         System.out.println("Exception while moving file: " + e.getMessage());
	      }
	      if(result != null) {
	         System.out.println("Tabelle wurde verschoben!");
	      }else{
	         System.out.println("File movement failed.");
	      }
		}

}
