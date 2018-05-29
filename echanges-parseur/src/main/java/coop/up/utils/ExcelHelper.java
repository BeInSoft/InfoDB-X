package coop.up.utils;

import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.NumberToTextConverter;



public class ExcelHelper {

	/**
	 * Méthode utilitaire, permet de retourner une chaine de caractères
	 * 
	 * @param cell
	 * @return
	 */
	public static String getStringFromCell(Cell cell) {
		String result = "";
		if (cell != null) {
			int cellType = cell.getCellType();
			switch (cellType) {
			case Cell.CELL_TYPE_BOOLEAN:
				result = cell.getBooleanCellValue() ? "true" : "false";
				break;
			case Cell.CELL_TYPE_NUMERIC:
				// quand c'est un numérique
				result = NumberToTextConverter.toText(cell.getNumericCellValue());
				break;
			case Cell.CELL_TYPE_STRING:
				// quand c'est une chaine de caractères
				result = cell.getStringCellValue();
				break;
			case Cell.CELL_TYPE_FORMULA:
				// quand c'est une chaine de caractères
				result = cell.getStringCellValue();
				break;
			case Cell.CELL_TYPE_BLANK: case Cell.CELL_TYPE_ERROR:
				result="";
				break;
			default:
				result="";
				break;
			}
		}
		return result;
	}
	
	public static String getFormatedDateFromCell(Cell cell, SimpleDateFormat formatter) {
		String result = "";
		if(cell!=null && cell.getDateCellValue()!=null) {
			result = formatter.format(cell.getDateCellValue());
		}
		return result;
	}

}
