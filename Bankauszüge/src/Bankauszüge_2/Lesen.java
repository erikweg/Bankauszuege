package Bankauszüge_2;

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;

public class Lesen {

	public static String inputpfad = "";
	
	public void setCSVpfad (String pfad) {
		inputpfad = pfad;
	}
	
	public String getWert(int zeile, int stelle) throws IOException {
		/*
		 * 'zeile' und 'stelle' sind 1 basiert
		 * 3. Zeile -> zeile = 3
		 * 5. Spalte -> spalte = 5
		 */
						
		stelle -= 1;
		String currentLine = null;
		FileReader fr = new FileReader(inputpfad);

		BufferedReader br = new BufferedReader(fr);

		for (int i = 0; i < zeile; i++) {
			currentLine = br.readLine();
		}
		String str[] = currentLine.split(";");

		for (int i = 0; i < str.length; i++) {
			str[i] = str[i].replaceAll("\"", ""); 			// Strings werden manchmal mit Anführungsstrichen gespeichert
		}

		br.close();

		return str[stelle];
	}
	
	public int getZeilenAnzahl() throws IOException {
				
		int anzahl = 0;
		String currentLine = null;
		FileReader fr = new FileReader(inputpfad);
		BufferedReader br = new BufferedReader(fr);
		while ((currentLine = br.readLine()) != null) {
			anzahl++;
		}
		
		br.close();
		return anzahl;
	}

	public double getBetrag(int zeile) throws IOException {
		
		String text = getWert(zeile, 15).replace(".", "").replace(",", ".");
		double betrag = Math.round(Double.valueOf(text) *100d) /100d;
		
		return betrag;
	}

	public String getDatumAnfang() throws IOException {
		
		String datum = this.getWert(this.getZeilenAnzahl(), 2);
		
		return datum;
	}
	
	public String getDatumEnde() throws IOException {
		
		String datum = this.getWert(2, 2);
		
		return datum;
	}

}
