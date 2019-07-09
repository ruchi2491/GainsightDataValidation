package com.atmecs.datavalidation;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Map.Entry;
import org.apache.log4j.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class FinalValidation {
	public ConstData condata;
	public static LinkedList<String> header;
	// logger for displaying statements
	static Logger logger = Logger.getLogger(FinalValidation.class.getName());

	// main method
	public static void main(String[] args) throws Exception {
		BasicConfigurator.configure();
		FinalValidation app = new FinalValidation();
		ConstData condata = new ConstData();
		try {
			condata.prop();
		} catch (Exception e2) {
			logger.error(e2.getMessage());
		}
		app.condata = condata;
		// logger.info("------"+condata.FILEPATH1);

		
		//  cretaing new folder
		  
		  
		  new File(condata.FILEPATH1 + "\\notrequiredfiles").mkdir(); new
		  File(condata.FILEPATH2 + "\\notrequiredfiles").mkdir();
		  
		 // convert csv to xlsx
		  
		  
		  ConvertCsvToExcel.convertCsvToXls(condata.FILEPATH1 + "\\");
		  ConvertCsvToExcel.convertCsvToXls(condata.FILEPATH2+ "\\");
		  
		  //move .csv file into another folder
		  
		  
		  ConvertCsvToExcel.movefile(condata.FILEPATH1);
		  ConvertCsvToExcel.movefile(condata.FILEPATH2);
		 

		// getFileList() will be call here for collecting list of sheets
		ArrayList<String> excelfiles1 = app.getFileList(condata.FILEPATH1);
		ArrayList<String> excelfiles2 = app.getFileList(condata.FILEPATH2);
		// getKeyRowData() will be call here for collecting list of keys according to
		// that filename
		ArrayList<String[]> keyColumName;
		try {
			keyColumName = app.getKeyRowData();
			for (String fileName1 : excelfiles1) {
				boolean flag = false;
				for (String fileName2 : excelfiles2) {
					// Comparing sandbox and production file names
					if (fileName1.compareTo(fileName2) == 0) {
						flag = true;
						for (String[] fileName : keyColumName) {
							// comparing the keycolumn name(getting from getKeyRowData()) in the sheet1_prod
							if (fileName1.compareTo(fileName[0]) == 0) {
								try {
									// creating 4 results files for storing results
									app.CreateResultFiles(fileName1);
									app.verifyExcelFile(fileName1, fileName2, fileName);

									logger.info(
											"-------------------------" + fileName1 + "---------------------------");
									logger.info(
											"-------------------------Results files are created---------------------------");

								} catch (IOException e) {
									logger.error(e.getMessage());
								}
								break;
							}
						}
						break;
					}
				}
				if (!flag) {
					// if flag is false that will come to here
					logger.error("File name not matched:" + fileName1);
				}
			}
		} catch (IOException e1) {
			logger.error(e1.getStackTrace());
		}

		Thread.sleep(4000);
		RowCount.genrateResultSheet();

		try {
			logger.info("Press ENTER key to exit(0)");
			System.in.read();
			System.exit(0);
		} catch (IOException e) {
			logger.error(e.getStackTrace());
		}
	}

	// this method is for getting keycolumn related to the sheet
	private ArrayList<String[]> getKeyRowData() throws IOException {
		// Storing all rows for keycolumns in this arraylist
		ArrayList<String[]> KeyArray = new ArrayList<String[]>();
		try {
			String keyfileName = getFileList(condata.KEYFILEPATH).get(0);
			String sheetName = keyfileName.substring(0, keyfileName.indexOf("."));
			Workbook keyWorkbook = readExcel(condata.KEYFILEPATH, keyfileName);
			Sheet Sheet = keyWorkbook.getSheet(sheetName);
			int rowCount = Sheet.getLastRowNum();
			logger.info("No.of rows in sheet regarding filenames and keycolumns---" + rowCount);
			for (int i = 1; i < rowCount + 1; i++) {
				Row row = Sheet.getRow(i);
				int noCell = row.getLastCellNum() + 1;
				String rowDataArray[] = new String[noCell];
				int j = 0;
				for (j = 0; j < noCell; j++) {
					if (row.getCell(j) != null) {
						rowDataArray[j] = row.getCell(j).toString();
					} else {
						rowDataArray[j] = "";
					}
				}
				KeyArray.add(rowDataArray);
			}
		} catch (Exception e) {
			logger.error(e.getMessage());
		}
		return KeyArray;

	}

	// This method is just for creating 4 results files in a folder based on sheet
	// name
	private void CreateResultFiles(String fileName) throws IOException {
		String folderName = fileName.substring(0, fileName.indexOf("."));
		File file = new File(condata.RESULTFILEPATH + "\\" + folderName);
		if (!file.exists()) {
			if (file.mkdir()) {
				logger.info("Directory is created!");
			} else {
				logger.warn("Failed to create directory!");
			}
		}
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Result");

		FileOutputStream out1 = new FileOutputStream(
				condata.RESULTFILEPATH + "\\" + folderName + "\\" + "Same_Data_In_Both_Sheets_" + folderName + ".xlsx");
		FileOutputStream out2 = new FileOutputStream(condata.RESULTFILEPATH + "\\" + folderName + "\\"
				+ "Same_ID_Differnt_Data_In_Both_Sheets_" + folderName + ".xlsx");
		FileOutputStream out3 = new FileOutputStream(condata.RESULTFILEPATH + "\\" + folderName + "\\"
				+ "Differnt_IDs_In_Sheet1_Production_" + folderName + ".xlsx");
		FileOutputStream out4 = new FileOutputStream(condata.RESULTFILEPATH + "\\" + folderName + "\\"
				+ "Differnt_IDs_In_Sheet2_Sandbox_" + folderName + ".xlsx");

		workbook.write(out1);
		workbook.write(out2);
		workbook.write(out3);
		workbook.write(out4);
		workbook.close();
		out1.close();
		out2.close();
		out3.close();
		out4.close();
	}

	public void verifyExcelFile(String fileName1, String fileName2, String[] keyColumName) throws IOException {
		Workbook workbook1 = readExcel(condata.FILEPATH1, fileName1);
		Workbook workbook2 = readExcel(condata.FILEPATH2, fileName2);

		Sheet Sheet1 = workbook1.getSheetAt(0);
		Sheet Sheet2 = workbook2.getSheetAt(0);

		int rowCount1 = Sheet1.getLastRowNum();

		int rowCount2 = Sheet2.getLastRowNum();

		HashSet<String> keysSheet1 = new HashSet<String>();
		HashSet<String> keysSheet2 = new HashSet<String>();
		SimpleDateFormat formatter = new SimpleDateFormat("MM-dd-yyyy");
		HashMap<String, LinkedList<String>> rowDataSheet1 = new HashMap<String, LinkedList<String>>();
		HashMap<String, LinkedList<String>> rowDataSheet2 = new HashMap<String, LinkedList<String>>();
		ArrayList<Integer> KeyCellNo = new ArrayList<Integer>();

		Iterator<Row> iterator = Sheet1.iterator();

		Row nextRow = iterator.next();
		Iterator<Cell> cellIterator = nextRow.cellIterator();
		header = new LinkedList<String>();
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();

			header.add(cell.toString());
			for (int i = 1; i < keyColumName.length; i++) {

				if (cell.toString().trim().compareTo(keyColumName[i].trim()) == 0) {
					KeyCellNo.add(cell.getColumnIndex());

					break;
				}

			}
		}

		// array1 type checking and getting ID's

		System.out.println(KeyCellNo.toString());

		for (int i = 1; i < rowCount1 + 1; i++) {

			Row row = Sheet1.getRow(i);
			String finalKey = "";
			for (int j = 0; j < KeyCellNo.size(); j++) {
				if (row.getCell(KeyCellNo.get(j)) != null) {

					finalKey = finalKey + row.getCell(KeyCellNo.get(j)).toString().trim();
				}
			}

			LinkedList<String> rowData = new LinkedList<String>();
			for (int j = 0; j < nextRow.getLastCellNum(); j++) {
				Cell cell = row.getCell(j);
				if (cell != null) {
					if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
						if (HSSFDateUtil.isCellDateFormatted(cell)) {

							rowData.add(String.valueOf(formatter.format(cell.getDateCellValue())));

						} else {
							rowData.add(String.format("%.0f", cell.getNumericCellValue()));
						}

					} else {
						rowData.add(cell.toString());
					}

				} else {
					rowData.add("");
				}

			}

			rowDataSheet1.put(finalKey.trim(), rowData);
			keysSheet1.add(finalKey.trim());

		}

		// array2 type checking and getting ID's
		for (int i = 1; i < rowCount2 + 1; i++) {

			Row row = Sheet2.getRow(i);
			String finalKey = "";
			for (int j = 0; j < KeyCellNo.size(); j++) {
				if (row.getCell(KeyCellNo.get(j)) != null) {
					finalKey = finalKey + row.getCell(KeyCellNo.get(j)).toString().trim();
				}
			}

			LinkedList<String> rowData = new LinkedList<String>();
			for (int j = 0; j < nextRow.getLastCellNum(); j++) {
				Cell cell = row.getCell(j);
				if (cell != null) {

					if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
						if (HSSFDateUtil.isCellDateFormatted(cell)) {

							rowData.add(String.valueOf(formatter.format(cell.getDateCellValue())));

						} else {
							rowData.add(String.format("%.0f", cell.getNumericCellValue()));
						}

					} else {
						rowData.add(cell.toString());
					}

				} else {
					rowData.add("");
				}

			}
			rowDataSheet2.put(finalKey.trim(), rowData);
			keysSheet2.add(finalKey.trim());
		}
		workbook1.close();
		workbook2.close();
		HashSet<String> SameKeysInBothSheet = new HashSet<String>(keysSheet1);
		HashSet<String> differntKeysInSheet1 = new HashSet<String>(keysSheet1);
		HashSet<String> differntKeysInSheet2 = new HashSet<String>(keysSheet2);
		SameKeysInBothSheet.retainAll(keysSheet2);
		differntKeysInSheet1.removeAll(keysSheet2);
		differntKeysInSheet2.removeAll(keysSheet1);
		System.out.println("Same Key Both :" + SameKeysInBothSheet.size() + "------Differnt Key Sheet1 :"
				+ differntKeysInSheet1.size() + "---------Differnt Key Sheet2 :" + differntKeysInSheet2.size());

		WriteExcelMatchId(fileName1, SameKeysInBothSheet, rowDataSheet1, rowDataSheet2, 1);
		WriteExcelNotMatchId(fileName1, differntKeysInSheet1, rowDataSheet1, 3);
		WriteExcelNotMatchId(fileName2, differntKeysInSheet2, rowDataSheet2, 4);
		SameKeysInBothSheet = null;
		differntKeysInSheet1 = null;
		differntKeysInSheet2 = null;
		keysSheet1 = null;
		keysSheet2 = null;
		rowDataSheet1 = null;
		rowDataSheet2 = null;

	}

	// Id's matched in both sheets
	private void WriteExcelMatchId(String fileName, Set<String> sameKeysInBothSheet,
			Map<String, LinkedList<String>> rowDataSheet1, Map<String, LinkedList<String>> rowDataSheet2, int j)
			throws IOException {
		Map<String, LinkedList<String>> sameData = new HashMap<String, LinkedList<String>>();
		Map<String, LinkedList<String>> DifferntData1 = new HashMap<String, LinkedList<String>>();
		Map<String, LinkedList<String>> DifferntData2 = new HashMap<String, LinkedList<String>>();
		Iterator<String> itrator = sameKeysInBothSheet.iterator();

		while (itrator.hasNext()) {
			String key = itrator.next();
			LinkedList<String> data1 = new LinkedList<String>();

			String value1 = "";
			String value2 = "";

			LinkedList<String> rowData1 = rowDataSheet1.get(key);
			LinkedList<String> rowData2 = rowDataSheet2.get(key);

			for (int i = 0; i < rowData1.size(); i++) {
				if (rowData1.get(i) != null) {
					value1 = rowData1.get(i);
				}
				if (rowData2.get(i) != null) {
					value2 = rowData2.get(i);
				}
				if (value1.compareTo(value2) == 0) {
					// Keeping ID's of data matched
					data1.add(value1);
				} else {
					// Keeping ID's of data not matched
					DifferntData1.put(key, rowData1);
					DifferntData2.put(key, rowData2);
					data1.clear();
					break;
				}

			}
			if (data1.size() != 0) {
				sameData.put(key, data1);
			}

		}
		logger.info("-----same keys containing same data :" + sameData.size()
				+ "----- same keys containing different data in production:" + DifferntData1.size()
				+ "--------- same keys containing different data in sandbox:" + DifferntData2.size() + "-----");
		AppendDataIntoExcel(fileName, sameData, 1);
		// AppendDataIntoExcel(fileName,DifferntData1,2);
		// AppendDataIntoExcel(fileName,DifferntData2,2);
		sortAndUniqueAppendDataIntoExcel(fileName, DifferntData1, DifferntData2);
		sameData = null;
		DifferntData1 = null;
		DifferntData2 = null;
	}

	private void sortAndUniqueAppendDataIntoExcel(String fileName, Map<String, LinkedList<String>> DifferntData1,
			Map<String, LinkedList<String>> DifferntData2) throws IOException {

		String folderName = fileName.substring(0, fileName.indexOf("."));
		File file = null;
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet spreadSheet = workbook.createSheet("Result");
		XSSFRow row;
		XSSFCell cell;

		int i = 0;

		file = new File(condata.RESULTFILEPATH + "\\" + folderName + "\\Same_ID_Differnt_Data_In_Both_Sheets_"
				+ folderName + ".xlsx");
		workbook = (XSSFWorkbook) readExcel(condata.RESULTFILEPATH + "\\" + folderName,
				"Same_ID_Differnt_Data_In_Both_Sheets_" + folderName + ".xlsx");
		spreadSheet = (XSSFSheet) workbook.getSheet(("Result"));
		row = spreadSheet.createRow(i);
		CellStyle style1 = workbook.createCellStyle();
		style1.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
		style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		XSSFFont xSSFFont = workbook.createFont();
		xSSFFont.setFontName(HSSFFont.FONT_ARIAL);
		xSSFFont.setFontHeightInPoints((short) 10);
		xSSFFont.setColor(HSSFColor.GREEN.index);
		style1.setFont(xSSFFont);

		CellStyle style2 = workbook.createCellStyle();
		style2.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
		style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		System.out.println(header.toString());
		for (int j = 0; j < header.size(); j++) {
			cell = row.createCell(j);

			cell.setCellValue(header.get(j).toString().trim());
		}
		int rowCount = spreadSheet.getLastRowNum() + 1;

		String value1 = "";
		String value2 = "";

		for (Entry<String, LinkedList<String>> entry : DifferntData1.entrySet()) {
			String key = entry.getKey();
			// do something with key and/or tab

			LinkedList<String> rowData1 = DifferntData1.get(key);
			LinkedList<String> rowData2 = DifferntData2.get(key);
			XSSFRow row1 = spreadSheet.createRow(rowCount++);
			XSSFRow row2 = spreadSheet.createRow(rowCount++);

			for (int j = 0; j < rowData1.size(); j++) {
				value1 = rowData1.get(j).toString().trim();
				value2 = rowData2.get(j).toString().trim();

				if (value1.compareTo(value2) != 0) {
					XSSFCell cell1 = row1.createCell(j);
					cell1.setCellValue(value1);
					XSSFCell cell2 = row2.createCell(j);
					cell2.setCellValue(value2);
					cell1.setCellStyle(style1);
					cell2.setCellStyle(style2);
				} else {
					XSSFCell cell1 = row1.createCell(j);
					cell1.setCellValue(value1);
					XSSFCell cell2 = row2.createCell(j);
					cell2.setCellValue(value2);

				}
			}

		}
		FileOutputStream output_file = new FileOutputStream(file); // Open FileOutputStream to write updates
		workbook.write(output_file);
		workbook.close();
		output_file.close();
	}

	// Id's not matched in both sheets means id's is there one among the sheet
	private void WriteExcelNotMatchId(String fileName, Set<String> differntKeysInSheet,
			Map<String, LinkedList<String>> rowDataSheet, int caseNo) throws IOException {
		LinkedList<String> data = new LinkedList<String>();
		Map<String, LinkedList<String>> DifferntData = new HashMap<String, LinkedList<String>>();
		Iterator<String> itrator = differntKeysInSheet.iterator();
		while (itrator.hasNext()) {
			String key = itrator.next();
			data = rowDataSheet.get(key);
			DifferntData.put(key, data);
		}
		// logger.info("--------- Differnt Data :"+DifferntData.size()+"-----");
		AppendDataIntoExcel(fileName, DifferntData, caseNo);
	}

	// Writing data into excel
	private void AppendDataIntoExcel(String fileName, Map<String, LinkedList<String>> rowDataSheet, int caseNo)
			throws IOException {
		String folderName = fileName.substring(0, fileName.indexOf("."));
		File file = null;
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet spreadSheet = workbook.createSheet("Result");
		XSSFRow row;
		XSSFCell cell;
		int i = 0;

		switch (caseNo) {
		case 1: // Same_Data_In_Both_Sheets
			file = new File(
					condata.RESULTFILEPATH + "\\" + folderName + "\\Same_Data_In_Both_Sheets_" + folderName + ".xlsx");
			spreadSheet = (XSSFSheet) workbook.getSheet(("Result"));
			// creating row for column names
			row = spreadSheet.createRow(i);
			// logger.info(header.toString());
			for (int j = 0; j < header.size(); j++) {
				cell = row.createCell(j);
				// Writing value into cell
				cell.setCellValue(header.get(j).toString().trim());
			}
			for (Entry<String, LinkedList<String>> entry : rowDataSheet.entrySet()) {
				String key = entry.getKey();
				row = spreadSheet.createRow(++i);
				LinkedList<String> data = rowDataSheet.get(key);
				for (int j = 0; j < data.size(); j++) {
					cell = row.createCell(j);
					cell.setCellValue(data.get(j).toString().trim());
				}
			}
			break;

		/*
		 * case 2: //Same_ID_Differnt_Data_In_Both_Sheets file = new
		 * File(condata.RESULTFILEPATH+"\\"+folderName+"\\
		 * Same_ID_Differnt_Data_In_Both_Sheets_"+folderName+".xlsx"); workbook =
		 * (XSSFWorkbook) readExcel(condata.RESULTFILEPATH+
		 * "\\"+folderName,"Same_ID_Differnt_Data_In_Both_Sheets_"+folderName+".xlsx");
		 * spreadSheet = (XSSFSheet) workbook.getSheet(("Result")); row =
		 * spreadSheet.createRow(i); //logger.info(header.toString()); for (int j = 0; j
		 * < header.size(); j++) { cell = row.createCell(j); //Writing value into cell
		 * cell.setCellValue(header.get(j).toString().trim()); } int rowCount =
		 * spreadSheet.getLastRowNum()+1; row = spreadSheet.createRow(rowCount); for
		 * (Entry<String, LinkedList<String>> entry : rowDataSheet.entrySet()) { String
		 * key = entry.getKey(); row = spreadSheet.createRow(rowCount++);
		 * LinkedList<String> data =rowDataSheet.get(key); for (int j = 0; j <
		 * data.size(); j++) { cell = row.createCell(j);
		 * cell.setCellValue(data.get(j).toString().trim()); } } break;
		 */
		case 3: // Differnt_IDs_In_Sheet1_Production
			file = new File(condata.RESULTFILEPATH + "\\" + folderName + "\\Differnt_IDs_In_Sheet1_Production_"
					+ folderName + ".xlsx");
			spreadSheet = (XSSFSheet) workbook.getSheet(("Result"));
			row = spreadSheet.createRow(i);
			// logger.info(header.toString());
			for (int j = 0; j < header.size(); j++) {
				cell = row.createCell(j);
				// Writing value into cell
				cell.setCellValue(header.get(j).toString().trim());
			}
			for (Entry<String, LinkedList<String>> entry : rowDataSheet.entrySet()) {
				String key = entry.getKey();
				row = spreadSheet.createRow(++i);
				LinkedList<String> data = rowDataSheet.get(key);
				for (int j = 0; j < data.size(); j++) {
					cell = row.createCell(j);
					cell.setCellValue(data.get(j).toString().trim());
				}
			}
			break;

		case 4: // Differnt_IDs_In_Sheet2_Sandbox
			file = new File(condata.RESULTFILEPATH + "\\" + folderName + "\\Differnt_IDs_In_Sheet2_Sandbox_"
					+ folderName + ".xlsx");
			spreadSheet = (XSSFSheet) workbook.getSheet(("Result"));
			row = spreadSheet.createRow(i);
			// logger.info(header.toString());
			for (int j = 0; j < header.size(); j++) {
				cell = row.createCell(j);
				// Writing value into cell
				cell.setCellValue(header.get(j).toString().trim());
			}
			for (Entry<String, LinkedList<String>> entry : rowDataSheet.entrySet()) {
				String key = entry.getKey();
				row = spreadSheet.createRow(++i);
				LinkedList<String> data = rowDataSheet.get(key);
				for (int j = 0; j < data.size(); j++) {
					cell = row.createCell(j);
					cell.setCellValue(data.get(j).toString().trim());
				}
			}
			break;

		}
		FileOutputStream output_file = new FileOutputStream(file); // Open FileOutputStream to write updates
		workbook.write(output_file);
		workbook.close();
		output_file.close();
	}

	// Reading sheet inside the workbook by its name
	private Workbook readExcel(String filePath, String fileName) throws IOException {
		File file = new File(filePath + "\\" + fileName);
		Workbook Workbook = null;
		try {
			FileInputStream files = new FileInputStream(file);
			String fileExtensionName = fileName.substring(fileName.indexOf("."));
			if (fileExtensionName.equals(".xlsx")) {
				Workbook = new XSSFWorkbook(files);
			} else if (fileExtensionName.equals(".xls")) {
				Workbook = new HSSFWorkbook(files);
			}
		} catch (Exception e) {
			logger.error(e.getMessage());
		}
		return Workbook;
	}

	// reading file names in a folder
	private ArrayList<String> getFileList(String filePath) {
		File folder = new File(filePath);
		ArrayList<String> ExcelFileList = new ArrayList<String>();
		if (folder.exists()) {
			File[] listOfFiles = folder.listFiles();

			for (int i = 0; i < listOfFiles.length; i++) {
				if (listOfFiles[i].isFile()) {
					ExcelFileList.add(listOfFiles[i].getName());
					logger.info("Files name----:" + listOfFiles[i].getName());
				}
			}

		} else {
			logger.error(
					"Files are not there in that folder-----Given path is wrong,please check Config.Properties file.");
		}
		return ExcelFileList;
	}

}
