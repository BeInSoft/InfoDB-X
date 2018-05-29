package coop.up.actions;

import static coop.up.utils.ExcelHelper.getStringFromCell;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.jdom2.Document;
import org.jdom2.Element;

public class VideosAction implements IAction {

	@Override
	public void execute(Sheet currentSheet, Document document) {
		Element rootElement = document.getRootElement();
		Element rootFragmentElement = new Element(currentSheet.getSheetName());
		
		int index = 1;
		Row currentRow = currentSheet.getRow(index++);
				
		while(currentRow != null) {
			Element videoElement = new Element("video");	
			videoElement.setAttribute("titre", getStringFromCell(currentRow.getCell(0)));
			videoElement.setAttribute("source", getStringFromCell(currentRow.getCell(1)));
			videoElement.setAttribute("icon", getStringFromCell(currentRow.getCell(2)));
			videoElement.setAttribute("vues", getStringFromCell(currentRow.getCell(3)));
			videoElement.setAttribute("auteur", getStringFromCell(currentRow.getCell(4)));
			
			rootFragmentElement.addContent(videoElement);
			currentRow = currentSheet.getRow(index++);
		}
		rootElement.addContent(rootFragmentElement);

	}

}
