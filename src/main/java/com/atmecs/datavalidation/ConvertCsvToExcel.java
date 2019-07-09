package com.atmecs.datavalidation;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;

import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.commons.io.FilenameUtils;

public class ConvertCsvToExcel {

	public static final String FILE_DELIMITER = ",(?=([^\"]*\"[^\"]*\")*[^\"]*$)";
	public static final String FILE_EXTN = ".xlsx";
	public static final String FILE_NAME = "EXCEL_DATA";
	private static Logger logger = Logger.getLogger(ConvertCsvToExcel.class);

	public static String convertCsvToXls(String csvFilePath) {
		SXSSFSheet sheet = null;
		Workbook workBook = null;
		BufferedReader fileReader = null;
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
					String Fileloc=csvFilePath+listOfFiles[j].getName();
					fileReader = new BufferedReader(new InputStreamReader(new FileInputStream(new File(Fileloc)), "UTF-16"));
					workBook = new SXSSFWorkbook();
					sheet = (SXSSFSheet) workBook.createSheet("Sheet");
					int rowNum = 0;
					String line="";
					
					logger.info("Creating New .Xls File From The Already Generated .Csv File");
					
					while ((line = fileReader.readLine()) != null) {
						//System.out.println(line);
						Row currentRow = sheet.createRow(rowNum++);
						String[] tokens = line.split(FILE_DELIMITER);
						int x = 0;
						for (String token : tokens) {
							String trial=token.trim();
							if(trial.contains("\"")) {
							trial=trial.replace('"', ' ').trim();
							}
							currentRow.createCell(x).setCellValue(trial);	
							x++;
						}	
					}
					
					File file = new File(Fileloc);
					String filename = FilenameUtils.removeExtension(file.getName());
					generatedXlsFilePath = file.getParent() + "\\" + filename  + FILE_EXTN;
					
					logger.info("The File Is Generated At The Following Location= " + generatedXlsFilePath);
					
					fileOutputStream = new FileOutputStream(generatedXlsFilePath.trim());
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
						fileReader.close();
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
	
	
	public static void main(String[] args) throws Exception {
		BasicConfigurator.configure();
		String fileLoc = "";
		ConstData condata = new ConstData();
		condata.prop();
		System.out.println(condata.CSVFILEPATH);
		fileLoc = ConvertCsvToExcel.convertCsvToXls(condata.CSVFILEPATH);
		System.out.println(fileLoc);
	}
}


