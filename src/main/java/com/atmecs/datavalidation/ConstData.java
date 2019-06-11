package com.atmecs.datavalidation;

import java.io.*;
import java.util.Properties;

class ConstData {
    public Properties prop;
	public String FILEPATH1;
	public String FILEPATH2;
	public String RESULTFILEPATH;
	public String KEYFILEPATH;
	public String CSVFILEPATH; 
	public ConstData() {
		FILEPATH1="";
		FILEPATH2="";
		RESULTFILEPATH="";
		KEYFILEPATH="";
		
	}
	public void  prop() throws Exception {
	String projectLocation=System.getProperty("user.dir"); //storing relative path in a var
	prop=new Properties();
	File f=new File(projectLocation+"/config.properties");
	FileReader fr=new FileReader(f);
	prop.load(fr);
	  FILEPATH1=prop.getProperty("Productionfilepath");
	FILEPATH2=prop.getProperty("Sandboxfilepath");
	 RESULTFILEPATH=prop.getProperty("Resultsfilepath");
	 KEYFILEPATH = prop.getProperty("Sheet_KeyColumnfilepath");
	 CSVFILEPATH= prop.getProperty("csvfilepath");
	}

}
