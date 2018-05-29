package coop.up.actions;

import static coop.up.utils.ExcelHelper.getFormatedDateFromCell;
import static coop.up.utils.ExcelHelper.getStringFromCell;

import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.jdom2.Document;
import org.jdom2.Element;

public class PlanningAction implements IAction {

	@Override
	public void execute(Sheet currentSheet, Document document) {
		Element rootElement = document.getRootElement();
		Element rootFragmentElement = new Element(currentSheet.getSheetName());

		int index = 1;
		Row currentRow = currentSheet.getRow(index++);
		String currentSlot = "";
		Element slot = null;
		
		SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd/MM/yyyy");
		SimpleDateFormat simpleHourFormat = new SimpleDateFormat("HH:mm");
				
		while(currentRow != null) {
			String currentRowSlot = getStringFromCell(currentRow.getCell(1));
			if(slot== null || !currentSlot.equals(currentRowSlot)) {
				slot = new Element("creneau");
				slot.setAttribute("jour",getStringFromCell(currentRow.getCell(0)));
				slot.setAttribute("slot",getStringFromCell(currentRow.getCell(1)));
				slot.setAttribute("date",getFormatedDateFromCell(currentRow.getCell(2), simpleDateFormat));
				rootFragmentElement.addContent(slot);
				currentSlot=currentRowSlot;
			}
			Element atelier = new Element("atelier");
			
			atelier.setAttribute("id",String.valueOf(currentRow.getRowNum()));
			atelier.setAttribute("type",getStringFromCell(currentRow.getCell(3)));
			atelier.setAttribute("titre",getStringFromCell(currentRow.getCell(4)));
			atelier.setAttribute("ssTitre",getStringFromCell(currentRow.getCell(5)));
			atelier.setAttribute("niveau",getStringFromCell(currentRow.getCell(6)));
			atelier.setAttribute("salle",getStringFromCell(currentRow.getCell(7)));
			atelier.setAttribute("urlSalle",getStringFromCell(currentRow.getCell(8)));
			atelier.setAttribute("heureDebut",getFormatedDateFromCell(currentRow.getCell(9), simpleHourFormat));
			atelier.setAttribute("heureFin",getFormatedDateFromCell(currentRow.getCell(10), simpleHourFormat));
			atelier.setAttribute("duree",getFormatedDateFromCell(currentRow.getCell(11), simpleHourFormat));
			
			Element desc = new Element("description");
			desc.setText(getStringFromCell(currentRow.getCell(12)));
			atelier.addContent(desc);
			
			slot.addContent(atelier);
			
			currentRow = currentSheet.getRow(index++);
		}
			
		
		rootElement.addContent(rootFragmentElement);
	}

}
