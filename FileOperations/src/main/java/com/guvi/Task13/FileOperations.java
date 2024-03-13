package com.guvi.Task13;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileOperations {
	
	public void writeExcel() throws FileNotFoundException, IOException {
		
		
		File fileobj = new File(System.getProperty("user.dir") + "\\WriteExcel.xlsx");
		FileOutputStream file = new FileOutputStream(fileobj);
		
		XSSFWorkbook book = new XSSFWorkbook();
		XSSFSheet sheet = book.createSheet("Sheet1");
		
		sheet.createRow(0).createCell(0).setCellValue("Name");
		sheet.getRow(0).createCell(1).setCellValue("Age");
		sheet.getRow(0).createCell(2).setCellValue("Email");
		
		sheet.createRow(1).createCell(0).setCellValue("John Doe");
		sheet.getRow(1).createCell(1).setCellValue("30");
		sheet.getRow(1).createCell(2).setCellValue("john@test.com");
		
		sheet.createRow(2).createCell(0).setCellValue("Jane Doe");
		sheet.getRow(2).createCell(1).setCellValue("28");
		sheet.getRow(2).createCell(2).setCellValue("john@test.com");
		
		sheet.createRow(3).createCell(0).setCellValue("Bob Smith");
		sheet.getRow(3).createCell(1).setCellValue("35");
		sheet.getRow(3).createCell(2).setCellValue("jacky@example.com");
		
		sheet.createRow(4).createCell(0).setCellValue("Swapnil");
		sheet.getRow(4).createCell(1).setCellValue("37");
		sheet.getRow(4).createCell(2).setCellValue("swapnil@example.com");
		
		book.write(file);
		
		book.close();
		file.close();
		
	}

	
	public void readExcel() throws FileNotFoundException, IOException
	{
		
		FileInputStream inputfile = new FileInputStream(System.getProperty("user.dir") + "\\WriteExcel.xlsx");
		
		XSSFWorkbook book = new XSSFWorkbook(inputfile);
		XSSFSheet sheet = book.getSheet("Sheet1");
		

		for(int i = 0 ; i <= sheet.getLastRowNum() ; i++)
		{
			for(int j = 0 ; j <= 2 ; j++) 
			{
				System.out.print(book.getSheet("Sheet1").getRow(i).getCell(j) + "  ");
			}
			System.out.println();
		}
		
		
		book.close();
		inputfile.close();
		
	}
	
	
	public static void main(String[] args) {
		
		FileOperations fileobj = new FileOperations();
		
		try {
			fileobj.writeExcel();
			
			fileobj.readExcel();
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			System.out.println(e);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			System.out.println(e);
		}

		
	}

}
