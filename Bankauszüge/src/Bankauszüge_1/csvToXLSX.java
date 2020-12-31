package Bankauszüge_1;

import java.io.FileOutputStream;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;

import java.awt.Desktop;
import java.io.*;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class csvToXLSX {

	public static void main(String[] args) {

		try {

			String csvdatei = "C:\\Users\\ewegenaer.WERKSTATT-AC\\Desktop\\JavaEingang\\bankauszug.csv";
			String tabellepfad = "C:\\Users\\ewegenaer.WERKSTATT-AC\\Desktop\\JavaZwischenspeicher";

			switch (JOptionPane.showConfirmDialog(null, "Ist dies die richtige Datei?:\n\n" + csvdatei)) {

				case JOptionPane.YES_OPTION:
					// nix
					break;
	
				case JOptionPane.NO_OPTION:
					JFileChooser chooser1 = new JFileChooser(System.getProperty("user.home") + "\\Desktop");
					FileNameExtensionFilter filter = new FileNameExtensionFilter("CSV", "csv");
					chooser1.setFileFilter(filter);
					chooser1.setAcceptAllFileFilterUsed(false);
					if (chooser1.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
						csvdatei = chooser1.getSelectedFile().getAbsolutePath();
					} else {
						System.out.println("Abgebrochen !!! (001)");
					}
					break;
	
				case JOptionPane.CANCEL_OPTION:
					JOptionPane.showConfirmDialog(null, "Abgebrochen! (002)", "Abbruch", JOptionPane.DEFAULT_OPTION);
					System.exit(0);
					break;
				}

			switch (JOptionPane.showConfirmDialog(null, "Ist dies der richtige Zielordner?:\n\n" + tabellepfad)) {

				case JOptionPane.YES_OPTION:
					// nix
					break;
	
				case JOptionPane.NO_OPTION:
					JFileChooser chooser2 = new JFileChooser(System.getProperty("user.home") + "\\Desktop");
					chooser2.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
					chooser2.setAcceptAllFileFilterUsed(false);
					if (chooser2.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
						tabellepfad = chooser2.getSelectedFile().getAbsolutePath();
					} else {
						System.out.println("Abgebrochen !!! (003)");
					}
					break;
	
				case JOptionPane.CANCEL_OPTION:
					JOptionPane.showConfirmDialog(null, "Abgebrochen! (004)", "Abbruch", JOptionPane.DEFAULT_OPTION);
					System.exit(0);
					break;
				}

			System.out.println("Eingang:\n" + csvdatei + "\nAusgang:\n" + tabellepfad);

			Workbook workBook = new XSSFWorkbook();
			Sheet sheet = workBook.createSheet("sheet1");
			String currentLine = null;
			int RowNum = -1; 																		// rowNum muss bei -1 anfangen, sonst ist erste Zeile leer
			BufferedReader br = new BufferedReader(new FileReader(csvdatei));
			while ((currentLine = br.readLine()) != null) {
				String str[] = currentLine.split(";");
				RowNum++;
				Row currentRow = sheet.createRow(RowNum);
				for (int i = 0; i < str.length; i++) {
					str[i] = str[i].replaceAll("\"", ""); 											// Strings werden manchmal mit Anführungsstrichen gespeichert
					currentRow.createCell(i).setCellValue(str[i]);
				}
			}

			FileOutputStream fileOutputStream = new FileOutputStream(tabellepfad + "\\bankauszugtabelle.xlsx");
			workBook.write(fileOutputStream);
			fileOutputStream.close();
			workBook.close();
			br.close();

			JOptionPane.showConfirmDialog(null, ".csv wurde in .xlsx umgewandelt !\n\n" + tabellepfad, "Erfolgreich",
					JOptionPane.DEFAULT_OPTION);

			Desktop desktop = Desktop.getDesktop();
			try {
				File datei = new File(tabellepfad);
				desktop.open(datei);
			} catch (IllegalArgumentException iae) {
				System.out.println("Fehler (005)");
			}

		} catch (Exception ex) {
			System.out.println(ex.getMessage() + "Exception in try");
		}

	}

}
