package Bankauszüge_1;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;

public class dateieingabe {

	public static void main(String[] args) {
		
		String csvdatei = "C:\\Users\\ewegenaer.WERKSTATT-AC\\Desktop\\JavaEingang\\bankauszug.csv";
		String tabelledatei = "C:\\Users\\ewegenaer.WERKSTATT-AC\\Desktop\\JavaZwischenspeicher\\bankauszugtabelle.xlsx";
		
		int antwortcsvdatei = JOptionPane.showConfirmDialog(null, "Ist dies die richtige Datei?:\n\n" + csvdatei);
		switch (antwortcsvdatei) {
		
		case JOptionPane.YES_OPTION:
			// nix
			break;
			
		case JOptionPane.NO_OPTION:
			JFileChooser chooser1 = new JFileChooser(System.getProperty("user.home") + "\\Desktop");
			FileNameExtensionFilter filter = new FileNameExtensionFilter("CSV", "csv");
			chooser1.setFileFilter(filter);
			chooser1.setAcceptAllFileFilterUsed(false);
			if(chooser1.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
				csvdatei = chooser1.getSelectedFile().getAbsolutePath();
			} else {
				System.out.println("Keine Auswahl !!! (Fehler 001)");
			}
			break;
			
		case JOptionPane.CANCEL_OPTION:
			JOptionPane.showConfirmDialog(null, "Abgebrochen!", "Abbruch", JOptionPane.DEFAULT_OPTION);
			System.exit(0);
			break;
		}
		
		int antworttabellenausgabe = JOptionPane.showConfirmDialog(null,  "Ist dies der richtige Zielordner?:\n\n" + tabelledatei);
		switch (antworttabellenausgabe) {
		
		case JOptionPane.YES_OPTION:
			//nix
			break;
		
		case JOptionPane.NO_OPTION:
			JFileChooser chooser2 = new JFileChooser(System.getProperty("user.home") + "\\Desktop");
			chooser2.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
			chooser2.setAcceptAllFileFilterUsed(false);
			if(chooser2.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
				tabelledatei = chooser2.getSelectedFile().getAbsolutePath();
			} else {
				System.out.println("Keine Auswahl !!! (Fehler 002)");
			}
			break;
			
		case JOptionPane.CANCEL_OPTION:
			JOptionPane.showConfirmDialog(null, "Abgebrochen!", "Abbruch", JOptionPane.DEFAULT_OPTION);
			System.exit(0);
			break;
		}
		
		System.out.println("Eingang:\n" + csvdatei + "\nAusgang:\n" + tabelledatei);
	}

}
