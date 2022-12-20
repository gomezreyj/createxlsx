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

import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.w3c.dom.Node;
import org.w3c.dom.Element;
import org.xml.sax.SAXException;

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
import javax.xml.parsers.ParserConfigurationException;

import org.apache.commons.cli.ParseException;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.Option;













import org.apache.commons.io.FileUtils;

import com.seu.encrypt.*;

public class CreateBigXlsx {
	
	public final static String DRIVER = "oracle.jdbc.driver.OracleDriver";
//	public final static String URL = "jdbc:oracle:thin:@(DESCRIPTION =(ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = any)(PORT = 1521))(LOAD_BALANCE = yes))(CONNECT_DATA =(SERVICE_NAME = gisphp03)))";
	public final static String URL = "jdbc:oracle:thin:";
	final static Logger logger = Logger.getLogger(CreateBigXlsx.class);
	
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
			logger.error("ERROR: Loading Oracle DRIVER");
			e.printStackTrace();
		}   
        Connection conn = null;
        try {
			conn = DriverManager.getConnection(URL + service,user,pass);
			logger.info("Connected to Oracle");
		} catch (SQLException e) {
			logger.error("ERROR: Getting connection to Oracle");
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
        	 logger.error("Private Key file not found.");
		 } 
         catch (IOException ioe) 
         {
        	logger.error("IO Exception");
		    ioe.printStackTrace();
		 } catch (ClassNotFoundException cnfe) {
			logger.error("Class Not Found");
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
	public String getJarFilesPath(String privateKey, String fileNameToDecrypt)
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
	        	  //System.out.println("JAR Entry: " + entry.getName());
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
				logger.error("IO Exception UnPacking JAR");
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
		} 
		catch (IOException e) 
		{
			logger.error("IO Exception Deleting Directory");
			e.printStackTrace();
		}
	}
	
	/*
	 * This method create a document using the report config file.
	 * Parameters:
	 *    fXmlFile: the name of the xml associated with the report
	 * Return: a Document object
	 */
	
	Document getDocument(File xmlFile)
	{	
		Document doc = null;
		DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
    	DocumentBuilder dBuilder;
		try {
			dBuilder = dbFactory.newDocumentBuilder();
			doc = dBuilder.parse(xmlFile);
	    	doc.getDocumentElement().normalize();
		} catch (ParserConfigurationException pe) {
			logger.error("PARSE: " + pe.getMessage());
		}
		catch( IOException ioe)
		{
			logger.error("I/O:" + ioe.getMessage());
		}
		catch(SAXException saxe)
		{
			logger.error("SAX: " + saxe.getMessage());
		}
    	return doc;
	}

	/*
	 * This method instantiate a transport object that contains
	 * all the report config values
	 * Parameters:
	 *    doc: the document associated with the report
	 * Return: all the config values from the xml report config
	 */
	ReportConfigTO instantiateAttributes(Document doc)
	{
		ReportConfigTO reportTO = new ReportConfigTO();
		
		reportTO.setUser(doc.getElementsByTagName("user").item(0).getTextContent());
		reportTO.setFiletodecrypt(doc.getElementsByTagName("filetodecrypt").item(0).getTextContent());
		logger.info("File to be Decrypted: " + reportTO.getFiletodecrypt());
		reportTO.setPk(doc.getElementsByTagName("pk").item(0).getTextContent());
    	logger.info("Private Key File Name: " + reportTO.getPk());  	
    	reportTO.setDescription(doc.getElementsByTagName("description").item(0).getTextContent());
    	reportTO.setXlsxname(doc.getElementsByTagName("xlsxname").item(0).getTextContent());
    	reportTO.setLogfile(doc.getElementsByTagName("logfile").item(0).getTextContent());	
		return reportTO;
	}
	
	/*
	 * This method get the properties of the log4j in order to find the name of the log file.
	 * Parameters:
	 *     logFile: name of the physical log4j property file
	 * Return: nothing 
	 */
	
    Properties setProperties(String logFile)
	{
		Properties props = new Properties();
    	InputStream in = this.getClass().getClassLoader().getResourceAsStream("com/se/uti/resources/log4j.properties");
    	try {
			props.load(in);
		} catch (IOException io) {
			logger.error("PROPERTY FILE: " + io.getMessage());
		}   	
    	logger.info("LOGFILENAME: " + props.getProperty("log4j.appender.file.File"));
    	
    	// Lets set dynamically the name of the log file
    	props.setProperty("log4j.appender.file.File", logFile); //reportTO.getLogfile());
    	//PropertiesConfigurator is used to configure logger from properties file
        PropertyConfigurator.configure(props); 
        return props;
	}
    
    XSSFCellStyle getDisclaimerStyle(SXSSFWorkbook wb)
    {
        XSSFCellStyle disclaimerStyle = (XSSFCellStyle) wb.createCellStyle();
	    disclaimerStyle.setWrapText(true);
	    disclaimerStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
	    disclaimerStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        return disclaimerStyle;
    }
    
    SXSSFSheet getDisclaimerSheet(SXSSFWorkbook wb, String disclaimer, XSSFCellStyle disclaimerStyle)
    {
    	SXSSFSheet disclaimerSheet = null; 
    	disclaimerSheet = (SXSSFSheet) wb.createSheet("DISCLAIMER");   
	    Row rowDisclaimer = disclaimerSheet.createRow((short) 1);
        Cell disclaimerCelValue = null;
        disclaimerCelValue = rowDisclaimer.createCell((short) 1);
        disclaimerCelValue.setCellValue(disclaimer);
        disclaimerCelValue.setCellStyle(disclaimerStyle);         
        disclaimerSheet.addMergedRegion(new CellRangeAddress(1,4,1,11));
        return disclaimerSheet;
    }
	  

    public static void main(String[] args) throws Throwable {
    	
    	PrivateKey privateK = null;
    	String query = null;
        ResultSet rs = null;
    	
        CreateBigXlsx bigxlsx = new CreateBigXlsx();
    	//File fXmlFile = new File("C:\\CreateXLSX_1.0\\createxlsx\\CreateBigXlsx\\projectqueue.xml");
    	File fXmlFile = new File(args[0]);
    	
    	Document doc = bigxlsx.getDocument(fXmlFile);
    	ReportConfigTO reportTO = bigxlsx.instantiateAttributes(doc);     
    	String pathToPk = bigxlsx.getJarFilesPath(reportTO.getPk(), reportTO.getFiletodecrypt());
    	
        Properties props = bigxlsx.setProperties(reportTO.getLogfile());                  
        // Lets get the private key
        privateK =  bigxlsx.getPrivateKeyFile(pathToPk + "\\" + reportTO.getPk());
        logger.info("Decrypting...");
        
        //Lets decrypt the passwd file
        byte[] dbpass =  EncryptionUtil.decrypt(new File(pathToPk + "\\" + reportTO.getFiletodecrypt()), privateK);
        String passwd = new String(dbpass);
        Connection conn = bigxlsx.getConnection(reportTO.getDescription(),reportTO.getUser(),passwd);
    	PreparedStatement stmt = null;
    	     	
    	logger.info("Creating WoorkBook...");
        SXSSFWorkbook wb = new SXSSFWorkbook(100); // keep 100 rows in memory, exceeding rows will be flushed to disk
        logger.info("Finished Creating WoorkBook");
                
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
            
         XSSFCellStyle disclaimerStyle = bigxlsx.getDisclaimerStyle(wb);  
      
         SXSSFSheet disclaimerSheet = null;
         String disclaimer = props.getProperty("disclaimer"); 
         disclaimerSheet =  bigxlsx.getDisclaimerSheet(wb, disclaimer, disclaimerStyle);
         logger.info("DISCLAIMER = " + disclaimer);               
                      
           // Iterate over the list of queries
           //Getting the query list   	
           Iterator<String> iter = bigxlsx.getListQueries(doc).iterator();
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
        	  logger.info("Preparing Statement");
              stmt = conn.prepareStatement(query);
              logger.info("Finished Preparing Statement");
              logger.info("Executing Query: " + query);
              rs = stmt.executeQuery();
              logger.info("Finished Executing Query");
              
              //Let get the metadata of the table
              ResultSetMetaData metaData = rs.getMetaData();
              int colCount = metaData.getColumnCount();
              logger.info("Number of Columns:" + colCount);
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
            logger.info("Populating Sheet...");
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
          logger.info("Finished Populating Sheet...");
        }     
        // Get the date of the report and append it to the filename
        DateFormat df = new SimpleDateFormat("yyyyMMdd");
        java.util.Date date =  new java.util.Date();
        
        FileOutputStream fileOut = new FileOutputStream(reportTO.getXlsxname() + "_" + df.format(date) + ".xlsx");     
        wb.write(fileOut);
        logger.info("Finished Creating Workbook...");
        conn.close();
        stmt.close();
        rs.close();
        fileOut.close();
        
        logger.info("Deleting Temporary Files..");
        bigxlsx.deleteDir("temp1");
        if ( wb.dispose() )
        	logger.info("Temporary Files Deleted Successfully...");
        else 
        	logger.info("ERROR: Could not deleted Temporary Files !!");
        
        logger.info("Closing Workbook");     
        wb.close();    
        logger.info("Process Finished Successfully");
    }
    
}
