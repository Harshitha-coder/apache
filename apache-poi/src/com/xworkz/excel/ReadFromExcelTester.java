package com.xworkz.excel;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.*;

public class ReadFromExcelTester {

	public static void main(String[] args) throws IOException {

		String excelFile = ".\\datafile\\PalaceInfo.xlsx";

		FileInputStream inputStream = new FileInputStream(excelFile);

		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

		XSSFSheet sheet = workbook.getSheetAt(0);

		Iterator itr = sheet.iterator();
		while (itr.hasNext()) {
			XSSFRow row = (XSSFRow) itr.next();
			Iterator cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				XSSFCell cell = (XSSFCell) cellIterator.next();
				switch (cell.getCellType()) {
				case STRING:
					System.out.print(cell.getStringCellValue());
					break;
				}
				System.out.print(" ");
			}
			System.out.println();
		}

	}

}
