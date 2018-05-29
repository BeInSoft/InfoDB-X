package coop.up;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Date;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.jdom2.Document;
import org.jdom2.Element;
import org.jdom2.output.Format;
import org.jdom2.output.XMLOutputter;

import coop.up.actions.IAction;

public class Main {

	// Définition des constantes
	private final static String DESKTOP_PATH = System.getProperty("user.home") + "/Desktop";
	private final static String INPUT_FILE = "/infodb/input.xlsx";
	private final static String TEMPLATE_FILE = "/infodb/template.xml";
	private final static String OUTPUT_FILE = "/infodb/output.xml";

	private final static String ROOT_ELEMENT = "infodb";

	public static void main(String[] args) {
		try {
			File template = new File(DESKTOP_PATH + TEMPLATE_FILE);
			File input = new File(DESKTOP_PATH + INPUT_FILE);
			Document document = null;

			if (template.exists()) {
				// TODO parser le fichier pour intégrer les éléments complémentaires

			} else {
				Element root = new Element(ROOT_ELEMENT);
				root.setAttribute("last-change", new Date().toString());
				document = new Document(root);
			}

			// Lecture du fichier Excel en entrée et parcours des feuilles du classeur
			final Workbook inputFile = WorkbookFactory.create(new FileInputStream(input));
			int numberOfSheets = inputFile.getNumberOfSheets();
			for (int i = 0; i < numberOfSheets; i++) {
				Sheet currentSheet = inputFile.getSheetAt(i);
				final String sheetName = currentSheet.getSheetName(); 
				
				//Pattern commande pour invoquer une action spécifique pour chaque feuille du classeur 
				Class<?> IActionClass = Class.forName("coop.up.actions."+sheetName.substring(0, 1).toUpperCase()+sheetName.substring(1)+"Action");
				IAction  currentAction = (IAction)IActionClass.newInstance();
				currentAction.execute(currentSheet, document);
			}

			XMLOutputter XMLOutput = new XMLOutputter(Format.getPrettyFormat());
			XMLOutput.output(document, new FileOutputStream(new File(DESKTOP_PATH + OUTPUT_FILE)));

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
