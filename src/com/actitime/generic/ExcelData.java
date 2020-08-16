package com.actitime.generic;

import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelData 
{
	//To get the data
	public static String getData(String filePath, String sheetName, int rn, int cn) 
	{
		try 
		{
//			FileInputStream fis = new FileInputStream(new File(filePath));
//			DataFormatter df = new DataFormatter();
//			Sheet sh = WorkbookFactory.create(fis).getSheet(sheetName);
//			String data = sh.getRow(rn).getCell(cn).getStringCellValue();
			
			FileInputStream fis = new FileInputStream(new File(filePath));
			DataFormatter df = new DataFormatter();
			Sheet sh = WorkbookFactory.create(fis).getSheet(sheetName);
			Cell cellData = sh.getRow(rn).getCell(cn);
			String data = df.formatCellValue(cellData);
			return data;
			
//			FileInputStream fis = new FileInputStream(filePath);
//			Workbook wb = WorkbookFactory.create(fis);
//			String data = wb.getSheet(sheetName).getRow(rn).getCell(cn).toString();
//			return data;
		}
		catch(Exception e) 
		{
			return "";
		}
	}
	
	//To get the row count
	public static int getRowCount(String filePath, String sheetName)
	{
		try 
		{
			FileInputStream fis = new FileInputStream(filePath);
			Workbook wb = WorkbookFactory.create(fis);
			int rc = wb.getSheet(sheetName).getLastRowNum();
			return rc;
		}
		catch(Exception e) 
		{
			return -1;
		}
	}
	
	//To get the cell count
	public static int getCellCount(String filePath, String sheetName, int rn)
	{
		try 
		{
			FileInputStream fis = new FileInputStream(filePath);
			Workbook wb = WorkbookFactory.create(fis);
			int cc = wb.getSheet(sheetName).getRow(rn).getLastCellNum();
			return cc;
		}
		catch(Exception e) 
		{
			return -1;
		}
	}
}
