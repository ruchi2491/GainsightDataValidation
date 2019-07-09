package com.atmecs.datavalidation;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class RowCount {
	static ConstData contdata = new ConstData();
	static File folder;
	static File folder2;
	static File[] listOfFiles;
	static File[] listOfFiles2;
	static String filename, filename2;
	static XSSFWorkbook readworkbook, readworkbook2, readresultworkbook, writeworkbook;
	static XSSFSheet sheet, sheet1, resultsheet;
	static String column_name;

	public static void main(String[] args) throws Exception {
		 RowCount.genrateResultSheet();
		//RowCount.getColouredcell();
	}

	public static void InitalizeVariable() {

	}

	public static void genrateResultSheet() throws Exception {

		System.out.println("In Rowcount");
		writeworkbook = new XSSFWorkbook();
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
		folder = new File(contdata.FILEPATH1);
		listOfFiles = folder.listFiles();

		folder2 = new File(contdata.FILEPATH2);
		listOfFiles2 = folder2.listFiles();

		// System.out.println(listOfFiles);

		for (int i = 0; i < listOfFiles.length; i++) {

			if (listOfFiles[i].isFile() && listOfFiles2[i].isFile()) {
				filename = listOfFiles[i].getName().substring(0, listOfFiles[i].getName().lastIndexOf("."));
				filename2 = listOfFiles2[i].getName().substring(0, listOfFiles2[i].getName().lastIndexOf("."));

				readworkbook = new XSSFWorkbook(listOfFiles[i]);
				sheet = readworkbook.getSheetAt(0);
				readworkbook2 = new XSSFWorkbook(listOfFiles2[i]);
				sheet1 = readworkbook2.getSheetAt(0);
				readresultworkbook = new XSSFWorkbook(new File(contdata.FILEPATH1).getParent() + "\\Results\\"
						+ filename + "\\Same_ID_Differnt_Data_In_Both_Sheets_" + filename + ".xlsx");
				//System.out.println(readresultworkbook.toString());
				resultsheet = readresultworkbook.getSheetAt(0);

				int dataChangedRecord = (resultsheet.getLastRowNum()) / 2;

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
					// System.out.println(sheet1.getLastRowNum());
					Cell cell2 = row.createCell(2);
					cell2.setCellValue(sheet1.getLastRowNum());
				}
				Cell cell3 = row.createCell(3);
				cell3.setCellValue(sheet1.getLastRowNum() - sheet.getLastRowNum());
				Cell cell4 = row.createCell(4);
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
			System.out.println("written successfully" + i);
		}
		System.out.println("!!!!!!!!!!!!!!!!!!written successfully!!!!!!!!!!!!!!!!!");

	}

	public static void getColouredcell() throws IOException {
		readresultworkbook = new XSSFWorkbook(
				"D:\\Gainsight\\24_8_2018\\2nd Deployment\\Results\\export_accounts\\Same_ID_Differnt_Data_In_Both_Sheets_export_accounts.xlsx");
		resultsheet = readresultworkbook.getSheetAt(0);
		boolean flag = false;
		ArrayList<String> column_names = new ArrayList<String>();
		// int totalcount = resultsheet.getLastRowNum();
			
		for (int i = 0; i < resultsheet.getLastRowNum(); i++) {
			Cell cell = resultsheet.getRow(i).getCell(4);
			CellStyle cellStyle = cell.getCellStyle();
			Color color = cellStyle.getFillForegroundColorColor();
			// System.out.println(cell.getAddress() + ": " + ((XSSFColor)
			// color).getARGBHex());
			if (color != null) {
				if (color instanceof XSSFColor) {
					flag = true;
					System.out.println(cell.getAddress() + ": " + ((XSSFColor) color).getARGBHex());
				} else if (color instanceof HSSFColor) {
					if (!(color instanceof HSSFColor.AUTOMATIC))
						System.out.println(cell.getAddress() + ": " + ((HSSFColor) color).getHexString());
				}
			} else {
				flag = false;
			}
			if (flag == true) {
				column_name = resultsheet.getRow(0).getCell(4).toString();
				System.out.println(column_name);
			}

		}

		// System.out.println();
		// for (Row row : resultsheet) {
		// for (Cell cell : row) {
		// if (! "".equals(String.valueOf(cell)))
		// System.out.println(cell.getAddress() + ": " + String.valueOf(cell));
		// CellStyle cellStyle = cell.getCellStyle();
		// Color color = cellStyle.getFillForegroundColorColor();
		// // System.out.println("color is:"+String.valueOf(color));
		// if (color != null) {
		// if (color instanceof XSSFColor) {
		// System.out.println(cell.getAddress() + ": " +
		// ((XSSFColor)color).getARGBHex());
		// } else if (color instanceof HSSFColor) {
		// if (! (color instanceof HSSFColor.AUTOMATIC))
		// System.out.println(cell.getAddress() + ": " +
		// ((HSSFColor)color).getHexString());
		// }
		// }
		// }
		// }
	}
}
