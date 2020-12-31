package Bankausz�ge_2;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class W�rter {
	
	String nix = "���������";
	DataFormatter fm = new DataFormatter();	
	String pfad = "";
	
	public void setW�rterPfad(String pfad) {
		this.pfad = pfad;
	}
	
	public Sheet createSheet() throws IOException {
		
		if(this.pfad.equals("")) {
			System.out.println("W�rter-Klasse hat keinen Pfad der XLSX-Datei f�r die Schlagw�rter erhalten. F�hre .setW�rterPfad() zuerst aus.");
			System.exit(0);
		}
		
		FileInputStream inputStream = new FileInputStream(new File(this.pfad));
		Workbook wb = new XSSFWorkbook(inputStream);
		Sheet sh = wb.getSheetAt(0);
		
		wb.close();
		inputStream.close();
		return sh;
	}
	
	public ArrayList<String> getNamenListe() throws IOException{
		
		ArrayList<String> namen = new ArrayList<>();
		Sheet sh = this.createSheet();
				
		int i = 1;
		boolean stop = false;
		String letzter = "���������";

		while(!stop){
			while(sh.getRow(i) != null && !stop){
				if(fm.formatCellValue(sh.getRow(i).getCell(0)).equals("ende")){
					stop = true;
				} else if(sh.getRow(i).getCell(0) == null || fm.formatCellValue(sh.getRow(i).getCell(0)).equals("")){
					namen.add(letzter);
				} else {
					namen.add(fm.formatCellValue(sh.getRow(i).getCell(0)));
					letzter = fm.formatCellValue(sh.getRow(i).getCell(0));
				}
			i++;
			}
		i++;
		}
		
		return namen;
	}
	
	public ArrayList<String> getSWVerwendungszweck() throws IOException{
		
		ArrayList<String> swverwendungszweck = new ArrayList<>();
		Sheet sh = this.createSheet();
		
		int i = 1;
		boolean stop = false;
				
		while(!stop) {
			while(sh.getRow(i) != null && !stop) {
				if(fm.formatCellValue(sh.getRow(i).getCell(0)).equals("ende")) {
					stop = true;
				} else if(sh.getRow(i).getCell(1) == null || fm.formatCellValue(sh.getRow(i).getCell(1)).equals("")) {
					swverwendungszweck.add(nix);
				} else {
					swverwendungszweck.add(fm.formatCellValue(sh.getRow(i).getCell(1)));
				}					
				
			i++;
			}
		i++;
		}		
		
		return swverwendungszweck;
	}
	
	public ArrayList<String> getSWBeg�nstigter() throws IOException{
		
		ArrayList<String> swbeg�nstigter = new ArrayList<>();
		Sheet sh = this.createSheet();
		
		int i = 1;
		boolean stop = false;
				
		while(!stop) {
			while(sh.getRow(i) != null && !stop) {
				if(fm.formatCellValue(sh.getRow(i).getCell(0)).equals("ende")) {
					stop = true;
				} else if(sh.getRow(i).getCell(2) == null || fm.formatCellValue(sh.getRow(i).getCell(2)).equals("")) {
					swbeg�nstigter.add(nix);
				} else {
					swbeg�nstigter.add(fm.formatCellValue(sh.getRow(i).getCell(2)));
				}					
				
			i++;
			}
		i++;
		}
		
		return swbeg�nstigter;
	}

	public ArrayList<Double> getLeereWerteListe(ArrayList<String> liste){
		
		ArrayList<Double> werte = new ArrayList<>();
		for(int i = 0; i < liste.size(); i++) {
			werte.add(0.0);
		}
		
		return werte;
	}
}
