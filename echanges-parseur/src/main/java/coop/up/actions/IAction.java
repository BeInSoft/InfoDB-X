package coop.up.actions;

import org.apache.poi.ss.usermodel.Sheet;
import org.jdom2.Document;

public interface IAction {

	void execute(Sheet currentSheet, Document document);
	
}
