package coop.up.actions;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.jdom2.Document;
import org.jdom2.Element;

import coop.up.utils.ExcelHelper;

public class TrombiAction implements IAction {

	@Override
	public void execute(Sheet currentSheet, Document document) {
		Element rootElement = document.getRootElement();
		Element rootFragmentElement = new Element(currentSheet.getSheetName());
		
		int index = 1;
		Row currentRow = currentSheet.getRow(index++);
		
		while(currentRow != null) {
			try {
				Element personne = new Element("personne");
				personne.setAttribute("id", String.valueOf(currentRow.getRowNum()));
				personne.setAttribute("civilite", ExcelHelper.getStringFromCell(currentRow.getCell(1)));
				personne.setAttribute("prenom", ExcelHelper.getStringFromCell(currentRow.getCell(2)));
				personne.setAttribute("nom", ExcelHelper.getStringFromCell(currentRow.getCell(3)));
				personne.setAttribute("telephone", ExcelHelper.getStringFromCell(currentRow.getCell(4)));
				personne.setAttribute("email", ExcelHelper.getStringFromCell(currentRow.getCell(5)));
				personne.setAttribute("direction", ExcelHelper.getStringFromCell(currentRow.getCell(6)));
				personne.setAttribute("metier", ExcelHelper.getStringFromCell(currentRow.getCell(7)));
				personne.setAttribute("equipe", ExcelHelper.getStringFromCell(currentRow.getCell(8)));;
				personne.setAttribute("entite", ExcelHelper.getStringFromCell(currentRow.getCell(9)));
				personne.setAttribute("bureau", ExcelHelper.getStringFromCell(currentRow.getCell(10)));
				
				//Ajout de la photo d'identité
				Element photo = new Element("photo");
				photo.setAttribute("source", ExcelHelper.getStringFromCell(currentRow.getCell(11)));
				personne.addContent(photo);
				
				for (int i = 0; i < 4; i++) {
					Element competence = new Element("competence");
					competence.setText(ExcelHelper.getStringFromCell(currentRow.getCell(12+i)));
					personne.addContent(competence);	
					
				}
				
				rootFragmentElement.addContent(personne);
				currentRow = currentSheet.getRow(index++);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		rootElement.addContent(rootFragmentElement);
		

	}

}
