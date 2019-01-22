package com.mycompany.app;

/**
 * Hello world!
 *
 */
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {

		// TODO Auto-generated method stub
		public static XSSFSheet ExcelWSheet;
		public static XSSFWorkbook ExcelWorkbook;
		public XSSFCell ExcelCell;
		public static XSSFRow ExcelRow;
		
		public static String filename_path="C:\\Users\\Manoj\\my-app\\Employee_Data.xlsx";
		
		public static void main(String args[]) throws IOException
		{
		FileInputStream ExcelFile = new FileInputStream(filename_path);
		ExcelWorkbook = new XSSFWorkbook(ExcelFile);
		ExcelWSheet = ExcelWorkbook.getSheet("Sheet1");
		int rowCount= ExcelWSheet.getLastRowNum() -ExcelWSheet.getFirstRowNum();
	
	
		// Create a loop for a reading row
		 for (int i = 0; i <=rowCount; i++) {

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

