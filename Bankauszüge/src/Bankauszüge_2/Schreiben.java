package Bankauszüge_2;

import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;

public class Schreiben {
	
	String pfad = "";
	
	public void setXLSXpfad(String pfad) {
		this.pfad = pfad;
	}
	
	public String getXLSXpfad() {
		return this.pfad;
	}
	
	public void createTabelle(Workbook wb, ArrayList<String> namen, ArrayList<Double> werte, double sonstige) throws IOException {
		
	// alle reihen und zellen die gebraucht werden könnten erschaffen
		
		Stuff st = new Stuff();
		int n = st.getAnzahlWerte(namen, werte);
		int z = n + 12;
		
		Sheet sh = wb.createSheet("Ausgaben");
		
		for(int i = 0; i < z; i++) {
			Row reihe = sh.createRow(i);
			reihe.setHeight((short) 400);
			for(int j = 0; j < 7; j++) {
				reihe.createCell(j);
			}
		}
		
	// datum und tage schreiben
		
		Lesen ls = new Lesen();
		
		sh.getRow(0).getCell(0).setCellValue("Anfangsdatum");
		sh.getRow(0).getCell(1).setCellValue(ls.getDatumAnfang());
		sh.getRow(1).getCell(0).setCellValue("Enddatum");
		sh.getRow(1).getCell(1).setCellValue(ls.getDatumEnde());
		sh.getRow(1).getCell(3).setCellValue("Tage");
		sh.getRow(1).getCell(4).setCellValue(st.getTage(ls.getDatumAnfang(),ls.getDatumEnde()));

	// ausgaben, einnahmen, summe
		
		sh.getRow(z - 5).getCell(0).setCellValue("Ausgaben:");
		sh.getRow(z - 5).getCell(3).setCellValue("Einnahmen:");
		sh.getRow(z - 1).getCell(3).setCellValue("Summe:");
		sh.getRow(z - 3).getCell(0).setCellValue("Sonstige:");
		//sh.getRow(6).getCell(0).setCellValue("Werte: " + st.getAnzahlWerte(namen, werte));
		//sh.getRow(7).getCell(0).setCellValue("z = " + z);
		
	// werte und namen		
		
		double ausgaben = 0;
		double einnahmen = 0;
		
		String letzter = "";
		int negz = 6;
		int posz = 6;
		
		for(int i = 0; i < namen.size(); i++) {
			
			if(!namen.get(i).equals(letzter) && werte.get(i) != 0) {
				
				if(werte.get(i) < 0) {
					ausgaben += werte.get(i);
					sh.getRow(negz).getCell(0).setCellValue(namen.get(i));
					sh.getRow(negz).getCell(1).setCellValue(werte.get(i));
					sh.getRow(negz).getCell(2).setCellValue("EUR");
					negz++;
				} else {
					einnahmen += werte.get(i);
					sh.getRow(posz).getCell(3).setCellValue(namen.get(i));
					sh.getRow(posz).getCell(4).setCellValue(werte.get(i));
					sh.getRow(posz).getCell(5).setCellValue("EUR");
					posz++;
				}					
			}
			
			letzter = namen.get(i);
		}
		
		sh.getRow(z - 5).getCell(1).setCellValue(ausgaben);
		sh.getRow(z - 5).getCell(2).setCellValue("EUR");
		sh.getRow(z - 5).getCell(4).setCellValue(einnahmen);
		sh.getRow(z - 5).getCell(5).setCellValue("EUR");
		sh.getRow(z - 3).getCell(1).setCellValue(sonstige);
		sh.getRow(z - 3).getCell(2).setCellValue("EUR");
		sh.getRow(z - 1).getCell(4).setCellValue(einnahmen + ausgaben + sonstige);
		sh.getRow(z - 1).getCell(5).setCellValue("EUR");

	// rahmen etc
		
		CellStyle ru = wb.createCellStyle();
		ru.setBorderBottom(BorderStyle.THICK);
		CellStyle rl = wb.createCellStyle();
		rl.setBorderLeft(BorderStyle.THICK);
		
		for(int i = 0; i < 6; i++) {
			sh.getRow(5).getCell(i).setCellStyle(ru);
		}
		
		for(int i = 6; i < (7 + n); i++) {
			sh.getRow(i).getCell(6).setCellStyle(rl);
		}
		
		for(int i = 0; i < 6; i++) {
			sh.getRow(6 + n).getCell(i).setCellStyle(ru);
		}
		
	// CellStyle für Summe	
		
		CellStyle sum = wb.createCellStyle();
		sum.setBorderTop(BorderStyle.DOUBLE);
		for(int i = 3; i < 6; i++) {
			sh.getRow(z - 1).getCell(i).setCellStyle(sum);
		}
		CellStyle sumcol = wb.createCellStyle();
		sumcol.setBorderTop(BorderStyle.DOUBLE);
		if(einnahmen + ausgaben + sonstige < 0) {
			sumcol.setFillForegroundColor(IndexedColors.RED.getIndex());
		} else {
			sumcol.setFillForegroundColor(IndexedColors.BRIGHT_GREEN1.getIndex());
		}
		sumcol.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		sh.getRow(z-1).getCell(4).setCellStyle(sumcol);
		
	// autosize	
		
		sh.autoSizeColumn(0);
		sh.setColumnWidth(1, 2500);
		sh.autoSizeColumn(3);
		sh.setColumnWidth(4, 2500);
		sh.setColumnWidth(2, 3000);
		sh.setColumnWidth(5, 3000);

	}
	
	public void createTabelleSonstige(Workbook wb, ArrayList<Integer> zeilen) throws IOException{
		
		Sheet sh = wb.createSheet("Sonstige");
		
		for(int i = 0; i < (zeilen.size() + 6); i++) {
			Row reihe = sh.createRow(i);
			reihe.setHeight((short) 600);
			for(int j = 0; j < 7; j++) {
				reihe.createCell(j);
			}
		}
		
		sh.getRow(1).getCell(0).setCellValue("Anzahl Einträge:");
		sh.getRow(1).getCell(1).setCellValue(zeilen.size());
		sh.getRow(3).getCell(1).setCellValue("Datum");
		sh.getRow(3).getCell(2).setCellValue("Verwendungszweck");
		sh.getRow(3).getCell(3).setCellValue("Begünstigter");
		sh.getRow(3).getCell(4).setCellValue("Betrag");
		
		Font fett = wb.createFont();
		fett.setBold(true);
		
		CellStyle t = wb.createCellStyle();
		t.setBorderBottom(BorderStyle.THICK);
		t.setAlignment(HorizontalAlignment.CENTER);
		t.setFont(fett);
		
		for(int i = 1; i < 6; i++) {
			sh.getRow(3).getCell(i).setCellStyle(t);
		}
		sh.getRow(3).setHeight((short) 600);
		
		Lesen ls = new Lesen();
		
		double sum = 0;
		
		for(int i = 0; i < zeilen.size(); i++) {
			
			sh.getRow(i + 4).getCell(1).setCellValue(ls.getWert(zeilen.get(i), 2));	// Datum
			
			sh.getRow(i + 4).getCell(2).setCellValue(ls.getWert(zeilen.get(i), 5));	// Verwendungszweck
			
			sh.getRow(i + 4).getCell(3).setCellValue(ls.getWert(zeilen.get(i), 12));// Begünstigter
			
			sh.getRow(i + 4).getCell(4).setCellValue(ls.getBetrag(zeilen.get(i)));	// Betrag
			
			sum += ls.getBetrag(zeilen.get(i));
			
			sh.getRow(i + 4).getCell(5).setCellValue("EUR");
			
		}
		
		CellStyle ft = wb.createCellStyle();
		ft.setBorderTop(BorderStyle.DOUBLE);
		ft.setFont(fett);
		
		sh.getRow(zeilen.size() + 5).getCell(4).setCellValue(sum);
		sh.getRow(zeilen.size() + 5).getCell(4).setCellStyle(ft);
		sh.getRow(zeilen.size() + 5).getCell(5).setCellValue("EUR");
		sh.getRow(zeilen.size() + 5).getCell(5).setCellStyle(ft);
		
		sh.setColumnWidth(1, 2500);
		sh.setColumnWidth(2, 15000);
		sh.setColumnWidth(3, 15000);
		sh.setColumnWidth(4, 2500);
		sh.setColumnWidth(5, 2500);		
		sh.autoSizeColumn(0);
		
	}
}
