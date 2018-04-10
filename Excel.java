package datadriven_framework;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel 
{
	static FileInputStream fi;
	static FileOutputStream fo;
	static XSSFWorkbook book;
	static XSSFSheet sht;
	static XSSFRow row;
	static XSSFCell cell;
	static String filepath="TestData\\";
	
	
	/*
	 * MethodName:--> Target workbook and sheet
	 */
	public static void Get_ExcelConnection(String filename, String sheetname) 
	{
		try {
			//Target file location
			fi=new FileInputStream(filepath+filename);
			book=new XSSFWorkbook(fi);
			sht=book.getSheet(sheetname);
			
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	
	/*
	 * MethodName:--> GetCellData
	 */
	public static String Get_CellData(int row, int cell)
	{
		return sht.getRow(row).getCell(cell).getStringCellValue();
	}
	
	
	
	/*
	 * MethodName:--> writeCellData
	 */
	public static void WriteCellData(int row, int cell, String result)
	{
		sht.getRow(row).createCell(cell).setCellValue(result);
	}
	

	/*
	 * MethodName:--> Create output file
	 */
	public static void Create_OPfile(String filename)
	{
		try {
			
			book.write(new FileOutputStream(filepath+filename));
			
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
