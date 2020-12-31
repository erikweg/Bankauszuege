package Bankauszüge_1;

import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class arrays_für_switchcase {

	public static void main(String[] args) {
		
		// Werte in Tabelle schreiben
		
				int n1x = 0; 
				int n2x = 0;
				
				for (int n1 = 0; n1 < namenlistekategorie.size(); n1++) {
					if (!namenlistekategorie.get(n1-1).equals(namenlistekategorie.get(n1)) || n1 == 0) {	
						Row reihe1 = sheet.createRow(n1x);
						reihe1.setHeight((short) 400);
						Cell zelle1 = reihe1.createCell(0);
						zelle1.setCellValue(namenlistekategorie.get(n1x));
						Cell zelle2 = reihe1.createCell(1);
						zelle2.setCellValue(wertelistekategorie.get(n1x) + "€");
						n1x = n1x + 1;
					} else {
						
					}
				}
				
				Row reihe3 = sheet.createRow(n1x + 1);
				reihe3.setHeight((short) 400);
				Cell zelle5 = reihe3.createCell(0);
				zelle5.setCellValue("Sonstiges");
				Cell zelle6 = reihe3.createCell(1);
				zelle6.setCellValue(kategoriesonstigewert);
				
				for (int n2 = 0; n2 < namenlistegruppe.size(); n2++) {
					int test001 = n1x + 3 + n2;
					System.out.println(test001);
					Row reihe2 = sheet.createRow(test001);
					reihe2.setHeight((short) 400);
					Cell zelle3 = reihe2.createCell(0);
					zelle3.setCellValue(namenlistegruppe.get(n2));
					Cell zelle4 = reihe2.createCell(1);
					zelle4.setCellValue(wertelistegruppe.get(n2) + "€");
					n2x = n2;
				}
				
				Row reihe4 = sheet.createRow(n1x + 3 + n2x + 1);
				reihe4.setHeight((short) 400);
				Cell zelle7 = reihe4.createCell(0);
				zelle7.setCellValue("Sonstiges");
				Cell zelle8 = reihe4.createCell(1);
				zelle8.setCellValue(gruppesonstigewert);
				
				
				sheet.autoSizeColumn(0);
				
		
		
		
		
		
		
		
	}

}
