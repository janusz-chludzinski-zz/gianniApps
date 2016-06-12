package pl.cadas.XMLPars;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XMLReader {
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	XSSFRow row;
	XSSFCell cell;
	
	File xlsxFile;
	
	public XMLReader (File file){
		this.xlsxFile = file;
	}
	
	public XSSFWorkbook readFile(){
		
		try {
			FileInputStream fis = new FileInputStream(xlsxFile);
			workbook = new XSSFWorkbook(fis);
			
			for(Sheet sheet : workbook){
				for(Row row : sheet){
					for(Cell cell : row){
						if(cell.getCellType() == Cell.CELL_TYPE_STRING){
							System.out.print(cell.getStringCellValue() + " ");
						} else {
							System.out.print(cell.getNumericCellValue() + " ");
						}
					}
					System.out.println();

				}
			}
			
		} catch (IOException e) {
			
			e.printStackTrace();
		}
		
		return workbook;
	}
	
}
