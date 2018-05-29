package coop.up.actions;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.jdom2.Document;
import org.jdom2.Element;

import static coop.up.utils.ExcelHelper.*;

public class OffresAction implements IAction {

	@Override
	public void execute(Sheet currentSheet, Document document) {
		Element rootElement = document.getRootElement();
		Element rootFragmentElement = new Element(currentSheet.getSheetName());
		
		int index = 1;
		Row currentRow = currentSheet.getRow(index++);
				
		while(currentRow != null) {
			Element offreElement = new Element("offre");	
			offreElement.setAttribute("nom", getStringFromCell(currentRow.getCell(0)));
			offreElement.setAttribute("fiche", getStringFromCell(currentRow.getCell(1)));
			offreElement.setAttribute("icon", getStringFromCell(currentRow.getCell(2)));
			
			//Ajout des images
			Element imagesElement = new Element("images");
			for (int i = 0; i < 5; i++) {
				String imagePath = getStringFromCell(currentRow.getCell(3+i));
				if(!"".equals(imagePath)) {
					Element imageElement = new Element("image");
					imageElement.setAttribute("source", imagePath);
					imagesElement.addContent(imageElement);
				}
			}
			offreElement.addContent(imagesElement);
			
			
			Element desc = new Element("description");
			desc.setText(getStringFromCell(currentRow.getCell(8)));
			offreElement.addContent(desc);
			
			rootFragmentElement.addContent(offreElement);
			currentRow = currentSheet.getRow(index++);
		}
		rootElement.addContent(rootFragmentElement);

	}

}
