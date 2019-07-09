package com.atmecs.datavalidation;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ResultGenration {

	@SuppressWarnings("resource")
	public static void main(String[] args) throws Exception {
		ConstData contdata = new ConstData();
		XSSFWorkbook readworkbook, readworkbook2, readresultworkbook;
		XSSFWorkbook writeworkbook = new XSSFWorkbook();
		XSSFSheet firstsheet = writeworkbook.createSheet("All record counts");
		XSSFRow row;

		XSSFCellStyle style_header = writeworkbook.createCellStyle();
		style_header.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
		style_header.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		XSSFCellStyle style_blue = writeworkbook.createCellStyle();
		style_blue.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());
		style_blue.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		XSSFCellStyle style_green = writeworkbook.createCellStyle();
		style_green.setFillForegroundColor(IndexedColors.LIME.getIndex());
		style_green.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		XSSFCellStyle style_orange = writeworkbook.createCellStyle();
		style_orange.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
		style_orange.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		contdata.prop();
		File folder = new File(contdata.FILEPATH1);
		File[] listOfFiles = folder.listFiles();

		File folder2 = new File(contdata.FILEPATH2);
		File[] listOfFiles2 = folder2.listFiles();

		// System.out.println(listOfFiles);

		for (int i = 0; i < listOfFiles.length; i++) {

			if (listOfFiles[i].isFile() && listOfFiles2[i].isFile()) {
				String filename = listOfFiles[i].getName().substring(0, listOfFiles[i].getName().lastIndexOf("."));
				String filename2 = listOfFiles2[i].getName().substring(0, listOfFiles2[i].getName().lastIndexOf("."));

				readworkbook = new XSSFWorkbook(listOfFiles[i]);
				XSSFSheet sheet = readworkbook.getSheetAt(0);
				readworkbook2 = new XSSFWorkbook(listOfFiles2[i]);
				XSSFSheet sheet1 = readworkbook2.getSheetAt(0);
				readresultworkbook = new XSSFWorkbook(new File(contdata.FILEPATH1).getParent() + "\\Results\\"
						+ filename + "\\Same_ID_Differnt_Data_In_Both_Sheets_" + filename + ".xlsx");
				XSSFSheet resultsheet=readresultworkbook.getSheetAt(0);
				
				int dataChangedRecord=(resultsheet.getLastRowNum())/2;
				
				XSSFRow headerrow = firstsheet.createRow(0);
				Cell headercell0 = headerrow.createCell(0);
				headercell0.setCellValue("Filename");
				Cell headercell1 = headerrow.createCell(1);
				headercell1.setCellValue("Rowcount of production");
				Cell headercell2 = headerrow.createCell(2);
				headercell2.setCellValue("Rowcount of sandbox");
				Cell headercell3 = headerrow.createCell(3);
				headercell3.setCellValue("Difference in count");
				Cell headercell4 = headerrow.createCell(4);
				headercell4.setCellValue("Data changed for records count");
				

				headerrow.getCell(0).setCellStyle(style_header);
				headerrow.getCell(1).setCellStyle(style_header);
				headerrow.getCell(2).setCellStyle(style_header);
				headerrow.getCell(3).setCellStyle(style_header);
				headerrow.getCell(4).setCellStyle(style_header);
				
				row = firstsheet.createRow(i + 1);
				Cell cell0 = row.createCell(0);
				cell0.setCellValue(filename);
				Cell cell1 = row.createCell(1);
				cell1.setCellValue(sheet.getLastRowNum());
				if (filename.equals(filename2)) {
					//System.out.println(sheet1.getLastRowNum());
					Cell cell2 = row.createCell(2);
					cell2.setCellValue(sheet1.getLastRowNum());
				}
				Cell cell3 = row.createCell(3);
				cell3.setCellValue(sheet1.getLastRowNum() - sheet.getLastRowNum());
				Cell cell4=row.createCell(4);
				cell4.setCellValue(dataChangedRecord);
				
				if (cell3.getNumericCellValue() > 0) {
					row.getCell(0).setCellStyle(style_blue);
					row.getCell(1).setCellStyle(style_blue);
					row.getCell(2).setCellStyle(style_blue);
					row.getCell(3).setCellStyle(style_blue);
					row.getCell(4).setCellStyle(style_blue);
				} else if (cell3.getNumericCellValue() == 0) {
					row.getCell(0).setCellStyle(style_orange);
					row.getCell(1).setCellStyle(style_orange);
					row.getCell(2).setCellStyle(style_orange);
					row.getCell(3).setCellStyle(style_orange);
					row.getCell(4).setCellStyle(style_orange);
				} else {
					row.getCell(0).setCellStyle(style_green);
					row.getCell(1).setCellStyle(style_green);
					row.getCell(2).setCellStyle(style_green);
					row.getCell(3).setCellStyle(style_green);
					row.getCell(4).setCellStyle(style_green);
				}

			} else {
				System.out.println("directory");
			}
			// Write the workbook in file system
			FileOutputStream out = new FileOutputStream(
					new File(new File(contdata.FILEPATH1).getParent() + "\\result.xlsx"));

			writeworkbook.write(out);
			out.close();
			System.out.println("written successfully");
		}
		System.out.println("!!!!!!!!!!!!!!!!!!written successfully!!!!!!!!!!!!!!!!!");
	}

}