package Bankausz¸ge_2;

import java.time.Duration;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;

public class Stuff {

	public String generateFilename(String datumanfang, String datumende, String path) {

		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("dd.MM.yy");
		LocalDate date1 = LocalDate.parse(datumanfang, dtf);
		LocalDate date2 = LocalDate.parse(datumende, dtf);

		DateTimeFormatter dtf2 = DateTimeFormatter.ofPattern("dd.LLL");
		DateTimeFormatter dtf3 = DateTimeFormatter.ofPattern("HHmmssSS");

		String name = path + "Auswertung " + dtf2.format(date1) + " bis " + dtf2.format(date2) + " "
				+ dtf3.format(LocalDateTime.now()) + ".xlsx";
		return name;
	}

	public double getTage(String datumanfang, String datumende) {
		
		DateTimeFormatter dtf = DateTimeFormatter.ofPattern("dd.MM.yy");
		LocalDate date1 = LocalDate.parse(datumanfang, dtf);
		LocalDate date2 = LocalDate.parse(datumende, dtf);
		
		double tage = Math.abs(Duration.between(date1.atStartOfDay(), date2.atStartOfDay()).toDays()) + 1;
		
		return tage;
	}
	
	public int getAnzahlWerte(ArrayList<String> namen, ArrayList<Double> werte) {

		int anzahl = 1;
		if (werte.get(0) == 0) {
			anzahl = 0;
		}

		int positive = 0;
		if (werte.get(0) > 0) {
			positive = 1;
		}

		for (int i = 1; i < namen.size(); i++) {
			if (namen.get(i) != namen.get(i - 1)) {
				if (werte.get(i) != 0) {
					anzahl++;
					if (werte.get(i) > 0) {
						positive++;
					}
				}
			}
		}

		int ges;
		if ((anzahl / 2) < positive) {
			ges = positive;
		} else {
			ges = anzahl - positive;
		}

		return ges;
	}

	public ArrayList<Double> addiereDoppelteWerte(ArrayList<String> namen, ArrayList<Double> werte){
		
		String letzter = "ƒ÷‹ƒ÷‹ƒ÷‹ƒ÷‹";
		
		int x = 0;
		
		for(int i = 0; i < namen.size(); i++) {
			
			if(namen.get(i).equals(letzter)) {
				werte.set((i - (1 + x)), Math.round((werte.get(i - (1 + x)) + werte.get(i)) *100d) /100d);
				x++;
			} else {
				x = 0;
			}
			
			letzter = namen.get(i);
		}
		
		return werte;
	}

	public double zeitMessen(double startTime, double endTime) {
		
		double execTime = Math.round(((endTime - startTime) / 1000000000) * 1000d) /1000d;
		
		return execTime;
	}
}
