package pl.cadas.XMLPars;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSWriter {

	private XSSFWorkbook XLSXWorkbook;
	private HSSFWorkbook outputWorkbook;
	private HSSFSheet outputSheet;
	private HSSFRow outputRow;
	private HSSFCell outputCell;
	private File file; 
	private FileOutputStream fos;
	
	public XLSWriter(XSSFWorkbook XLSXWorkbook){
		this.XLSXWorkbook = XLSXWorkbook;
	}
	
	public void writeXmlFile(){
		outputWorkbook = new HSSFWorkbook();
		outputSheet = outputWorkbook.createSheet();
		XSSFSheet XLSXsheet = XLSXWorkbook.getSheetAt(0);
		int rowCount = 0;
		
		for(Row row : XLSXsheet){
			
			if(rowCount == 100){
				outputSheet = outputWorkbook.createSheet();
				rowCount = 0;
			}
			
			outputRow = outputSheet.createRow(rowCount);
			rowCount++;
			
			for(Cell cell : row){
				
				if(cell.getCellType() == Cell.CELL_TYPE_STRING){
					outputCell = outputRow.createCell(cell.getColumnIndex());
					outputCell.setCellType(Cell.CELL_TYPE_STRING);
					
					outputCell.setCellValue(cell.getStringCellValue());
					
				} else {
					outputCell = outputRow.createCell(cell.getColumnIndex());
					outputCell.setCellType(Cell.CELL_TYPE_NUMERIC);
					
					outputCell.setCellValue(cell.getNumericCellValue());
				}
			}
		}
		
		file = new File("../XMLPars/outputXLS.xls");
		
		try {
			fos = new FileOutputStream(file);
			outputWorkbook.write(fos);
			fos.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
}
