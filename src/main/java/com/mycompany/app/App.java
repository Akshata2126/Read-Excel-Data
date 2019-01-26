package com.mycompany.app;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {

		// Variable to access ExcelFile
		public static XSSFSheet ExcelWSheet;
		public static XSSFWorkbook ExcelWorkbook;
		public XSSFCell ExcelCell;
		public static XSSFRow ExcelRow;
		
		//Location of Excel File
		public static String filename_path="C:\\Employee_Data.xlsx";
		
		public static void main(String args[]) throws IOException
		{
		//Create an object of FileInputStream class to read excel file
		FileInputStream ExcelFile = new FileInputStream(filename_path);
			
		//create object of XSSFWorkbook class
		ExcelWorkbook = new XSSFWorkbook(ExcelFile);
			
		//Selecting Sheet1 from Workbook
		ExcelWSheet = ExcelWorkbook.getSheet("Sheet1");
			
		//Find number of rows in excel sheet
		int rowCount= ExcelWSheet.getLastRowNum() -ExcelWSheet.getFirstRowNum();
	
		// Create a loop to read a row
		 for (int i = 0; i <=rowCount; i++) {
			
			 //Read a row
			 ExcelRow = ExcelWSheet.getRow(i);

		        //Create a loop to print cell values in a row
		        for (int j = 0; j < ExcelRow.getLastCellNum(); j++) {

		            //Print Excel data in console
		            System.out.print(ExcelRow.getCell(j)+" | ");
		        }

		        System.out.println();
		 }
			
		}

	}

