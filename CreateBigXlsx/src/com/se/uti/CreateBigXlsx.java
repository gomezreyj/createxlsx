package com.se.uti;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.math.BigDecimal;
import java.net.JarURLConnection;
import java.net.URI;
import java.net.URLConnection;
import java.security.PrivateKey;
import java.sql.Connection;
import java.sql.Date;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

import junit.framework.Assert;
import oracle.sql.JAVA_STRUCT;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.w3c.dom.Node;
import org.w3c.dom.Element;

import java.util.ArrayList;
import java.util.Enumeration;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Properties;
import java.util.jar.JarEntry;
import java.util.jar.JarFile;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.commons.cli.ParseException;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.Option;













import org.apache.commons.io.FileUtils;

import com.seu.encrypt.*;

public class CreateBigXlsx {
	
	public final static String DRIVER = "oracle.jdbc.driver.OracleDriver";
//	public final static String URL = "jdbc:oracle:thin:@(DESCRIPTION =(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = gisp01-crs.sempra.com)(PORT = 1521))(LOAD_BALANCE = yes))(CONNECT_DATA =(SERVICE_NAME = gisphp03)))";
	public final static String URL = "jdbc:oracle:thin:";
	
	/**
	   * String to hold the name of the private key file.
	   */
	public static final String PRIVATE_KEY_FILE = "irpk";

	static Properties props;
	
	/*
	 This method look to connect to Oracle DB.
	 Parameters:
	      service = the description of the service string ( host, port )
	      user =  username registered in the DB
	      pass = password of the user
	 Returns:
	      An connection object
	*/
	public Connection getConnection(String service, String user, String pass)
	{
		//Connect to Oracle
		//step1 load the driver class   
        try {
			Class.forName(DRIVER);
		} catch (ClassNotFoundException e) {
			System.out.println("ERROR: Loading Oracle DRIVER");
			e.printStackTrace();
		}   
        Connection conn = null;
        try {
			conn = DriverManager.getConnection(URL + service,user,pass);
        	System.out.println("Connected to Oracle");
		} catch (SQLException e) {
			System.out.println("ERROR: Getting connection to Oracle");
			e.printStackTrace();
		}   
        return conn;
	}
	
	/*
	  This method create a list of queries specified in the XML description file.
	  Parameters:
	     doc = The xml document that contains the views used to extract the data
	  Returns:
	     A list of queries 
	 * 
	 */
	public List<String> getListQueries(Document doc)
	{
		List<String> listQueries = new ArrayList<String>();
		NodeList nList = doc.getElementsByTagName("view");
    	
    	for (int temp = 0; temp < nList.getLength(); temp++) 
    	{ 
    		Node nNode = nList.item(temp);
    		listQueries.add(nNode.getTextContent());
    	}
		return listQueries;
	}
		
	/*
	 This method create a PrivateKey object taking as input the private key file then deserialize the file.
	 Parameters:
	     pk = the private key file. This file is referred using an absolute path to the temp umpacked JAR file
	 Returns:
	    A PrivateKey object 
	 */
	public PrivateKey getPrivateKeyFile(String pk)
	{
		 ObjectInputStream inputStream = null;
		 PrivateKey privateKey = null;
		 // getting the path to the private key file
		 
         try
         {         
		     inputStream = new ObjectInputStream(new FileInputStream(pk));
		     privateKey = (PrivateKey) inputStream.readObject();
		     inputStream.close();
         }
         catch(FileNotFoundException fnf)
		 {
			  System.out.println("Private Key file not found.");
		 } 
         catch (IOException ioe) 
         {
		    ioe.printStackTrace();
		 } catch (ClassNotFoundException cnfe) {
			cnfe.printStackTrace();
		}
        
         return privateKey;
         
	 }
	/*
	 This method unpack the JAR file and create a temporary directory that will hold the private and  password encrypted files.
	 The purpose of the unpacking is find the temporary absolute path where are located private and password file.
	 Parameters:
	    privateKey = name of the private key file
	    fileNameToDecrypt = name of the password file to be decrypted
	 Return:
	   The absolute path location for private key and encrypted password files
	 */
	public String unPackJar(String privateKey, String fileNameToDecrypt)
	{
		File nf = null;
		try
		{
		    File fj =new File("encviewsreport.jar");
	        JarFile jf;
			
	        jf = new JarFile(fj);
			Enumeration<JarEntry> entries = jf.entries();
	        String destdir = "temp1";
	        
	          while (entries.hasMoreElements())
	          {
	        	  JarEntry entry=(JarEntry)entries.nextElement();
	        	  if ( entry.getName().toLowerCase().contains(privateKey) || entry.getName().toLowerCase().contains(fileNameToDecrypt))
	        	  {
	        		 nf = new File(destdir,entry.getName());
	        		 
	        		 if(!nf.exists())
	                 {
	                    nf.getParentFile().mkdirs();
	                    nf = new java.io.File(destdir, entry.getName());
	                 }
	        		 java.io.InputStream is;
					 is = jf.getInputStream(entry);
					 java.io.FileOutputStream fos;
					 fos = new java.io.FileOutputStream(nf);
					
	        		 while (is.available() > 0) {  // write contents of 'is' to 'fos'
						    fos.write(is.read());
					 }
	        	     fos.close();
					 is.close();
	        	 }
	           }
	           jf.close();
	        }
			catch (IOException e) {
				e.printStackTrace();
			}
	        
	        String temp1Path = nf.getAbsolutePath().substring(0,nf.getAbsolutePath().lastIndexOf('\\'));
	        return temp1Path;
		
	}
	
	/*
	 * This method delete a directory from the current file system using Apache IO Library
	 * Parameters:
	 *   dirName = the name of the directory to delete
	 */
	
	public void deleteDir(String dirName)
	{
		try {
			FileUtils.deleteDirectory(new File(dirName));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	

    public static void main(String[] args) throws Throwable {
    	
    	PrivateKey privateK = null;
    	//Get the XML with the parameters to create the report
    	//File fXmlFile = new File("C:\\ExportReports\\projectqueue_enc.xml");
    	File fXmlFile = new File(args[0]);
        //File fXmlFile = new File("C:\\TestBigXLSX\\CreateBigXlsx\\phmsa.xml");
    	DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
    	DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
    	Document doc = dBuilder.parse(fXmlFile);
    	
    	doc.getDocumentElement().normalize();
    	
    	CreateBigXlsx bigxlsx = new CreateBigXlsx();
    	String usr = doc.getElementsByTagName("user").item(0).getTextContent();
    	
    	// Lets get the file to decrypt from the XML property file
    	String fileNameToDecrypt = doc.getElementsByTagName("filetodecrypt").item(0).getTextContent();
    	System.out.println(fileNameToDecrypt);
    	// Lets get the file to decrypt from the XML property file
    	String privateKey = doc.getElementsByTagName("pk").item(0).getTextContent();
    	System.out.println(privateKey);  	
 
    	String pathToPk = bigxlsx.unPackJar(privateKey, fileNameToDecrypt);       
    	//String pass = doc.getElementsByTagName("pass").item(0).getTextContent();
    	String desc = doc.getElementsByTagName("description").item(0).getTextContent();
    	String outPutFileName = doc.getElementsByTagName("xlsxname").item(0).getTextContent();
    	
    	//Getting the query list   	
        Iterator<String> iter = bigxlsx.getListQueries(doc).iterator();
                
        // Lets get the private key
        privateK =  bigxlsx.getPrivateKeyFile(pathToPk + "\\" + privateKey);
        System.out.println("before decrypt");
        //Lets decrypt the passwd file
        byte[] dbpass =  EncryptionUtil.decrypt(new File(pathToPk + "\\" + fileNameToDecrypt), privateK);
        String passwd = new String(dbpass);
        Connection conn = bigxlsx.getConnection(desc,usr,passwd);
    	PreparedStatement stmt = null;
    	     	
    	System.out.println("Creating WoorkBook");
        SXSSFWorkbook wb = new SXSSFWorkbook(100); // keep 100 rows in memory, exceeding rows will be flushed to disk
        System.out.println("Finished Creating WoorkBook");
                
         //Cell style for header row
           CellStyle cStyle = wb.createCellStyle();
           cStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
           cStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
           Font f = wb.createFont();
           f.setFontHeightInPoints((short) 12);
           cStyle.setFont(f);
           
           CellStyle dateStyle = wb.createCellStyle();
           DataFormat dateFormat = wb.createDataFormat();
           dateStyle.setDataFormat(dateFormat.getFormat("MM/dd/yyyy"));
            
           String query = null;
           ResultSet rs = null;
                      
           // Iterate over the list of queries
           while ( iter.hasNext())
           { 
        	  //New Sheet
              SXSSFSheet sheet = null;
          	  // Row and column indexes
              int idx = 0;
              int idy = 0;
              // Getting the query
        	  query = "SELECT * FROM " + iter.next();
        	  //Create the sheet
        	  sheet = (SXSSFSheet) wb.createSheet(query.substring(14));
        	  
        	  System.out.println("Preparing Statement");
            
        	  stmt = conn.prepareStatement(query);
              System.out.println("Finished Preparing Statement");
              System.out.println("Executing Query: " + query);
              rs = stmt.executeQuery();
              System.out.println("Finished Executing Query");
              
              //Let get the metadata of the table
              ResultSetMetaData metaData = rs.getMetaData();
              int colCount = metaData.getColumnCount();
              System.out.println("Number of Columns:" + colCount);
              int colType = 1;
              while ( colType <= colCount) {
                String columnType = metaData.getColumnClassName(colType);
                String columnName = metaData.getColumnName(colType);
                //System.out.println(columnType + "--" + columnName + "--" + metaData.getPrecision(colType));
                colType++;
              }
              
            //Create Hash Map of Field Definitions
              LinkedHashMap<Integer, TableInfo> hashMap = new LinkedHashMap<Integer, TableInfo>(colCount);
              for (int i = 0; i < colCount; i++) {
              	TableInfo ti = new TableInfo();
              	ti.setColumnType(metaData.getColumnClassName(i + 1));
              	ti.setColumnName(metaData.getColumnName(i + 1));
              	ti.setFieldSize(metaData.getPrecision(i + 1));
                hashMap.put(i,ti);
              }
              
              // Generate column headings
              Row row = sheet.createRow(idx);
                            
              TableInfo tableInfo = new TableInfo();
              Cell celValue = null;
              
              Iterator<Integer> iterator = hashMap.keySet().iterator();
              while (iterator.hasNext()) 
              {
                  Integer key = (Integer) iterator.next();
                  tableInfo = hashMap.get(key); 
                  celValue = row.createCell(idy);
                  celValue.setCellStyle(cStyle);
                  celValue.setCellValue(tableInfo.getColumnName());
                 
                  if(tableInfo.getFieldSize() > tableInfo.getColumnName().trim().length()){
                      sheet.setColumnWidth(idy, (tableInfo.getFieldSize() * 220 ) > 65025 ? 65025 : tableInfo.getFieldSize() * 220);
                  }
                  else {
                      sheet.setColumnWidth(idy, (tableInfo.getColumnName().trim().length() * 490 ) > 65025 ? 65025 :  tableInfo.getColumnName().trim().length() * 490 );
                  }
                  idy++;
              }
              
              
             System.out.println("Populating Sheet...");
            // Lets iterate over the result set and create a row per record
            // then create cells as much columns then populate each cell
            while ( rs.next() )
            {
              idx++;
              row = sheet.createRow(idx);
            
        	  for ( int idxCol = 1; idxCol <= colCount; idxCol++)
        	  {
        		celValue = row.createCell(idxCol - 1);
        		TableInfo t = hashMap.get(idxCol - 1);
				if ( t.getColumnType().equals("java.lang.String"))
				{
        			//System.out.println("String: " + (String)rs.getObject(idxCol));
        			if ( (String)rs.getObject(idxCol) != null)
        			{
				       celValue.setCellValue((String)rs.getObject(idxCol));
				       //System.out.println("String in cell:" + (String)rs.getObject(idxCol));
        			}
				}
				else
			    if ( t.getColumnType().equals("java.math.BigDecimal"))
			    {
	        	    //System.out.println("Number: " + (BigDecimal)rs.getObject(idxCol));
	        	    if ( (BigDecimal)rs.getObject(idxCol) != null)
	        	    {
	        	      BigDecimal bd = (BigDecimal) rs.getObject(idxCol);
	        	      Double dv = bd.doubleValue();
	        	      //System.out.println("Double in cell:" + dv);
	        	      celValue.setCellValue(dv);
	        	    }
			    }
			    else
				if ( t.getColumnType().equals("java.sql.Timestamp"))
				{
	        	    //System.out.println("Date: " + (java.sql.Timestamp)rs.getObject(idxCol));
	        	    if ( (java.sql.Timestamp)rs.getObject(idxCol) != null)
	        	    {
	        	     Timestamp stamp = (java.sql.Timestamp)rs.getObject(idxCol); 
	        	     Date date = new Date(stamp.getTime());
	        	     celValue.setCellStyle(dateStyle);
	        	     //System.out.println("Date in cell:" + date);
	        	     celValue.setCellValue(date);

	        	    }
				}
        	}
        	
          }
          System.out.println("Finished Populating Sheet...");
        }     
        // Get the date of the report and append it to the filename
        DateFormat df = new SimpleDateFormat("yyyyMMdd");
        java.util.Date date =  new java.util.Date();
        
        FileOutputStream fileOut = new FileOutputStream(outPutFileName + "_" + df.format(date) + ".xlsx");
        
        //FileOutputStream fileOut = new FileOutputStream(props.getProperty("excel_path").trim() + excelFilename.trim() );
        
        wb.write(fileOut);
        System.out.println("Finished Creating Workbook...");
        conn.close();
        stmt.close();
        rs.close();
        fileOut.close();
        
        System.out.println("Deleting Temporary Files..");
        bigxlsx.deleteDir("temp1");
        if ( wb.dispose() )
        	System.out.println("Temporary Files Deleted Successfully...");
        else 
        	System.out.println("ERROR: Could not deleted Temporary Files !!");
               
        wb.close();    

    }
    
}