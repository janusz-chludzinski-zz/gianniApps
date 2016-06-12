package pl.cadas.XMLPars;

import java.io.File;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
    	File xlsxFile = new File("../XMLPars/XLSXFile_1736kb.xlsx");
    	XMLReader reader = new XMLReader(xlsxFile);
    	XSSFWorkbook XLSXWorkbook = reader.readFile();
    	XLSWriter writer = new XLSWriter(XLSXWorkbook);
    	writer.writeXmlFile();
    }
}
