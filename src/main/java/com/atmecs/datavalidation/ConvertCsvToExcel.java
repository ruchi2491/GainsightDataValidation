package com.atmecs.datavalidation;

import com.opencsv.CSVReader;
import org.apache.commons.io.FilenameUtils;
import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.ArrayList;

public class ConvertCsvToExcel {

	public static final char FILE_DELIMITER = ',';
	public static final String FILE_EXTN = ".xlsx";
	public static final String FILE_NAME = "EXCEL_DATA";
	private static Logger logger = Logger.getLogger(ConvertCsvToExcel.class);

	public static String convertCsvToXls(String csvFilePath) {
		SXSSFSheet sheet = null;
		CSVReader reader = null;
		Workbook workBook = null;
		String generatedXlsFilePath = "";
		FileOutputStream fileOutputStream = null;

		File folder = new File(csvFilePath);
		File[] listOfFiles = folder.listFiles();
		
		ArrayList<String> notrequired=new ArrayList<String>();
		notrequired.add("export_accounts_contacts.csv");
		notrequired.add("export_cloud_trail_entitlements.csv");
		notrequired.add("export_xd_usage_rt.csv");
		notrequired.add("export_totango_manual_attributes_account.csv");
		notrequired.add("export_totango_manual_attributes_product.csv");
		
		
		for (int j = 0; j < listOfFiles.length; j++) {
			if (listOfFiles[j].isFile()) {
				System.out.println("File " + listOfFiles[j].getName());
				try {

					/**** Get the CSVReader Instance & Specify The Delimiter To Be Used ****/
					
					
					String Fileloc=csvFilePath+listOfFiles[j].getName();
					
					String[] nextLine;
					reader = new CSVReader(new FileReader(Fileloc), FILE_DELIMITER);
					workBook = new SXSSFWorkbook();
					sheet = (SXSSFSheet) workBook.createSheet("Sheet");
					int rowNum = 0;
					logger.info("Creating New .Xls File From The Already Generated .Csv File");
					while ((nextLine = reader.readNext()) != null) {
						Row currentRow = sheet.createRow(rowNum++);
						for (int i = 0; i < nextLine.length; i++) {
							System.out.println(nextLine[i]);
							currentRow.createCell(i).setCellValue(nextLine[i]);
						}
					}
					File file = new File(Fileloc);
					String filename = FilenameUtils.removeExtension(file.getName());
					generatedXlsFilePath = file.getParent() + "\\" + filename + FILE_EXTN;
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
						/**** Closing The CSV File-ReaderObject ****/
						reader.close();
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
		fileLoc = ConvertCsvToExcel.convertCsvToXls(condata.CSVFILEPATH);
		System.out.println(fileLoc);
	}

	public static void movefile(String csvfileloc) {
		File folder = new File(csvfileloc);
		File[] listOfFiles = folder.listFiles();

		for (int j = 0; j < listOfFiles.length; j++) {
			String ext1 = FilenameUtils.getExtension(listOfFiles[j].toString());
			System.out.println(ext1);
			if (ext1.equals("csv")) {
				 if(listOfFiles[j].renameTo
				           (new File(csvfileloc + "\\notrequiredfiles\\" + listOfFiles[j].getName())))
				        {
				            // if file copied successfully then delete the original file
					 listOfFiles[j].delete();
				            System.out.println("File moved successfully");
				        }
				        else
				        {
				            System.out.println("Failed to move the file");
				        }
			}
		}

	}
}
