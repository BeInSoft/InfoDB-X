package coop.up.actions;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.jdom2.Document;
import org.jdom2.Element;

import static coop.up.utils.ExcelHelper.*;

public class RetroAction implements IAction {

	@Override
	public void execute(Sheet currentSheet, Document document) {
		Element rootElement = document.getRootElement();
		Element rootFragmentElement = new Element(currentSheet.getSheetName());
		
		int index = 1;
		Row currentRow = currentSheet.getRow(index++);
		Element currentEdition = null;
		String currentEditionYear = "";
		
				
		while(currentRow != null) {
			String currentYear = NumberToTextConverter.toText(currentRow.getCell(0).getNumericCellValue());
			if(currentEdition==null || !currentEditionYear.equals(currentYear)) {
				if(currentEdition!=null) {
					rootFragmentElement.addContent(currentEdition);
				}
				currentEdition = new Element("edition");
				currentEdition.setAttribute("annee", currentYear);
				currentEdition.setAttribute("lieu", getStringFromCell(currentRow.getCell(1)));
				currentEdition.setAttribute("nombreDePhotos", getStringFromCell(currentRow.getCell(2)));
				currentEditionYear=currentYear;
			}
			Element image = new Element("image");
			image.setAttribute("source",getStringFromCell(currentRow.getCell(3)));
			currentEdition.addContent(image);
			
	
			currentRow = currentSheet.getRow(index++);
		}
		if(currentEdition!=null) {
			rootFragmentElement.addContent(currentEdition);
		}
		rootElement.addContent(rootFragmentElement);
		
	}

}
