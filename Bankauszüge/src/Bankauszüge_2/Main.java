package Bankauszüge_2;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	public static void main(String[] args) throws IOException {
		
		double startTime = System.nanoTime();
		
		Lesen ls = new Lesen();
		Fenster f = new Fenster();
		Wörter w = new Wörter();
		Schreiben sch = new Schreiben();
		Stuff st = new Stuff();
		
		double time1 = System.nanoTime();
		
		ls.setCSVpfad(f.pfadAbfrage("csv"));
		w.setWörterPfad(f.pfadAbfrage("xlsx"));
		sch.setXLSXpfad(f.pfadAbfrage("dir"));
		
		double time2 = System.nanoTime();
		double usertime = st.zeitMessen(time1, time2);		
		
		ArrayList<String> namen = w.getNamenListe();
		ArrayList<String> swverwendungszweck = w.getSWVerwendungszweck();
		ArrayList<String> swbegünstigter = w.getSWBegünstigter();
		ArrayList<Double> werte = w.getLeereWerteListe(namen);
		
		double sonstsum = 0;
		ArrayList<Integer> sonstzeilen = new ArrayList<>();

		double time3 = System.nanoTime();
		System.out.println("ArrayLists erstellt: Zeit gebraucht: " + st.zeitMessen(time2, time3) + " s");
		
		int n = ls.getZeilenAnzahl();
		
		for (int j = 2; j <= n; j++) {						// j = 1 weil erste Zeile Überschriften sind
			
			double time3_1 = System.nanoTime();
			
			int kategorie = -1; 												// kategorie == -1 ist 'sonstige'
			
			for (int i = 0; i < swverwendungszweck.size(); i++) {
				if (ls.getWert(j, 5).contains(swverwendungszweck.get(i))) {
					kategorie = i;
				}
			}
			
			for (int i = 0; i < swbegünstigter.size(); i++) {
				if (ls.getWert(j, 12).contains(swbegünstigter.get(i))) {
					kategorie = i;
				}
			}
			
			if(kategorie > -1) {
				werte.set(kategorie, (Math.round((werte.get(kategorie) + ls.getBetrag(j)) *100d) /100d));
			} else {
				sonstsum = Math.round((sonstsum + ls.getBetrag(j)) *100d) / 100d;
				sonstzeilen.add(j);
			}
			
			double time3_2 = System.nanoTime();
			int prozent = (j *100 / n);
			System.out.println("\t" + j + "/" + n + ", " + prozent + "%, " + st.zeitMessen(time3_1, time3_2) + " s");
		}
		
		double time4 = System.nanoTime();
		System.out.println("Daten ausgewertet: Zeit gebraucht: " + st.zeitMessen(time3, time4) + " s");		
		
		Workbook wb = new XSSFWorkbook();
		
		werte = st.addiereDoppelteWerte(namen, werte);
		
		sch.createTabelle(wb, namen, werte, sonstsum);
		sch.createTabelleSonstige(wb, sonstzeilen);
		
		try {
			FileOutputStream output = new FileOutputStream(st.generateFilename(ls.getDatumAnfang(), ls.getDatumEnde(), sch.getXLSXpfad()));
			wb.write(output);
			output.close();
		} catch(Exception e) {
			e.printStackTrace();
		}
		
		wb.close();
		
		double endTime = System.nanoTime();
		System.out.println("Datei gespeichert: Zeit gebraucht: " + st.zeitMessen(time4, endTime) + " s");
		System.out.println("Fertig! Execution Time: " + (st.zeitMessen(startTime, endTime) - usertime) + " s");
		
	}

}
