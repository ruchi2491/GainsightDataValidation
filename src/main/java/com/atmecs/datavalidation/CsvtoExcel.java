package com.atmecs.datavalidation;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Scanner;

import org.apache.commons.io.FilenameUtils;
import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.opencsv.CSVReader;

public class CsvtoExcel {

	public static final String FILE_DELIMITER = ",";
	public static final String FILE_EXTN = ".xlsx";
	public static final String FILE_NAME = "EXCEL_DATA";
	private static Logger logger = Logger.getLogger(ConvertCsvToExcel.class);

	public static String convertCsvToXls(String csvFilePath) {
		SXSSFSheet sheet = null;
		BufferedReader fileReader = null;
		Workbook workBook = null;
		String generatedXlsFilePath = "";
		FileOutputStream fileOutputStream = null;

		File folder = new File(csvFilePath);
		File[] listOfFiles = folder.listFiles();

		for (int j = 0; j < listOfFiles.length; j++) {
			if (listOfFiles[j].isFile()) {
				System.out.println("File " + listOfFiles[j].getName());
				try {
					String Fileloc = csvFilePath + listOfFiles[j].getName();
					workBook = new SXSSFWorkbook();
					sheet = (SXSSFSheet) workBook.createSheet("Sheet");
					logger.info("Creating New .Xls File From The Already Generated .Csv File");

					String line = "";
					// Create the file reader
					fileReader = new BufferedReader(new InputStreamReader(new FileInputStream(new File(Fileloc)), "UTF-16"));
					String specialChars = "/*!@#$%^&*()\"{}_[]|\\?/<>,.";
					int rowNum = 0;
					// Read the file line by line
					while ((line = fileReader.readLine()) != null) {
						System.out.println(line);
						Row currentRow = sheet.createRow(rowNum++);
						// Get all tokens available in line
						String[] tokens = line.split(FILE_DELIMITER);
						int x = 0;
						for (String token : tokens) {
							currentRow.createCell(x).setCellValue(token);	
							x++;
						}	
					}

					File file = new File(Fileloc);
					String filename = FilenameUtils.removeExtension(file.getName());
					generatedXlsFilePath = file.getParent() + "\\" + filename + "1" + FILE_EXTN;
					logger.info("The File Is Generated At The Following Location= " + generatedXlsFilePath);
					fileOutputStream = new FileOutputStream(generatedXlsFilePath.trim());
					System.out.println(fileOutputStream);
					workBook.write(fileOutputStream);
				} catch (Exception exObj) {
					logger.error("Exception In convertCsvToXls() Method?=  " + exObj);
				} finally {
					try {
						/**** Closing The Excel Workbook Object ****/
						workBook.close();
						/**** Closing The File-Writer Object ****/
						fileOutputStream.close();

					} catch (IOException ioExObj) {
						logger.error("Exception While Closing I/O Objects In convertCsvToXls() Method?=  " + ioExObj);
					}
				}

			} else if (listOfFiles[j].isDirectory()) {
				System.out.println("Directory " + listOfFiles[j].getName());
			}
		}
		return generatedXlsFilePath;
	}

	public static void main(String[] args) throws Exception {
		BasicConfigurator.configure();
		String fileLoc = "";
		ConstData condata = new ConstData();
		condata.prop();
		System.out.println(condata.CSVFILEPATH);
		fileLoc = CsvtoExcel.convertCsvToXls(condata.CSVFILEPATH);
		System.out.println(fileLoc);
	}

}
