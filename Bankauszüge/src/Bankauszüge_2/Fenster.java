package Bankauszüge_2;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;

public class Fenster {
	
	public String pfadAbfrage(String dateiformat) {
		
		String pfad = "";
		String format = "";
		String beschr = "";
		String frage = "";
		boolean filterbool = true;
		
		switch (dateiformat) {
		case "csv":
			pfad = "C:\\Users\\erikw\\Desktop\\JavaEingang\\bankauszug_groß.csv";
			format = "csv";
			beschr = "CSV";
			frage = "Ist dies die richtige CSV-Datei des Bankauszugs?\n\n";
			break;
		
		case "xlsx":
			pfad = "C:\\Users\\erikw\\Desktop\\JavaEingang\\Wörter.xlsx";
			format = "xlsx";
			beschr = "XLSX";
			frage = "Ist dies die richtige XSLX-Datei mit den Schlagwörtern?\n\n";
			break;
			
		case "dir":
			pfad = "C:\\Users\\erikw\\Desktop\\JavaZiel\\";
			format = "dir";
			beschr = "Verzeichnis";
			frage = "Ist dies das richtige Verzeichnis um die Auswertung abzuspeichern?\n\n";
			filterbool = false;
			break;
			
		default:
			System.out.println("Falsches Dateiformat in Fenster.pfadAbfrage()! [csv / xlsx / dir]");
			System.exit(0);
		}
		
		switch (JOptionPane.showConfirmDialog(null, frage + pfad)) {
		
		case JOptionPane.YES_OPTION:
			// nix
			break;
			
		case JOptionPane.NO_OPTION:
			JFileChooser chooser = new JFileChooser(System.getProperty("user.home") + "\\Desktop");
			
			if(filterbool) {
				FileNameExtensionFilter filter = new FileNameExtensionFilter(beschr, format);
				chooser.setFileFilter(filter);
			} else {
				chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
			}
			chooser.setAcceptAllFileFilterUsed(false);
			
			if(chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
				pfad = chooser.getSelectedFile().getAbsolutePath();
			} else {
				System.out.println("Abgebrochen! (001)");
			}
			break;
			
		case JOptionPane.CANCEL_OPTION:
			JOptionPane.showConfirmDialog(null, "Abgebrochen! (002)", "Abbruch", JOptionPane.DEFAULT_OPTION, JOptionPane.ERROR_MESSAGE);
			System.exit(0);
			break;
		}
		
		
		return pfad;
	}
	
}
