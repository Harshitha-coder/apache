package com.xworkz.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CreateAndWriteExcelTester {

	public static void main(String[] args) throws IOException {
		
		XSSFWorkbook workbook=new XSSFWorkbook();
		XSSFSheet sheet=workbook.createSheet("Task");
		sheet.createRow(0);
		sheet.getRow(0).createCell(0).setCellValue("Name");
		sheet.getRow(0).createCell(1).setCellValue("Language");
		sheet.getRow(0).createCell(2).setCellValue("Popular");
		sheet.getRow(0).createCell(3).setCellValue("NoOfVillans");
		
		sheet.createRow(1);
		sheet.getRow(1).createCell(0).setCellValue("Yuvarathna");
		sheet.getRow(1).createCell(1).setCellValue("Kannada");
		sheet.getRow(1).createCell(2).setCellValue(true);
		sheet.getRow(1).createCell(3).setCellValue(5);
		
		sheet.createRow(2);
		sheet.getRow(2).createCell(0).setCellValue("Colour Photo");
		sheet.getRow(2).createCell(1).setCellValue("Telugu");
		sheet.getRow(2).createCell(2).setCellValue(false);
		sheet.getRow(2).createCell(3).setCellValue(1);
		
		sheet.createRow(3);
		sheet.getRow(3).createCell(0).setCellValue("Dia");
		sheet.getRow(3).createCell(1).setCellValue("Kannada");
		sheet.getRow(3).createCell(2).setCellValue(true);
		sheet.getRow(3).createCell(3).setCellValue(0);
		
		sheet.createRow(4);
		sheet.getRow(4).createCell(0).setCellValue("Jathirathnalu");
		sheet.getRow(4).createCell(1).setCellValue("Telugu");
		sheet.getRow(4).createCell(2).setCellValue(true);
		sheet.getRow(4).createCell(3).setCellValue(2);
		
		sheet.createRow(5);
		sheet.getRow(5).createCell(0).setCellValue("Alidu Ulidavaru");
		sheet.getRow(5).createCell(1).setCellValue("Kannada");
		sheet.getRow(5).createCell(2).setCellValue(false);
		sheet.getRow(5).createCell(3).setCellValue(2);
		
		File file=new File("C:\\Users\\lenovo\\eclipse-workspace\\apache-poi\\ExcelFiles\\Test.xlsx");
		FileOutputStream fos=new FileOutputStream(file);
		workbook.write(fos);
		workbook.close();
		
	}

}
