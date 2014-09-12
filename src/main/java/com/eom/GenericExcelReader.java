package com.eom;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import com.eom.annotation.ExcelColumn;
import com.eom.handler.ExcelSheetContentsHandler;

public class GenericExcelReader<T extends IExcelEntity> {

	private static final Log logger = LogFactory.getLog(GenericExcelReader.class);

	public final static String EXCEL_TYPE_XLS = "xls";
	public final static String EXCEL_TYPE_XLSX = "xlsx";

	private File file;
	private final String fileType;

	private OPCPackage xlsxPackage;
	private ExcelSheetContentsHandler<T> sheetContentsHandler;

	/* --- Constructors --- */

	public GenericExcelReader(String filePath) {
		super();

		if(filePath == null || filePath.isEmpty())
			throw new IllegalArgumentException("Invalid filepath : empty or null");
		

		// Set the file
		this.file = new File(filePath);
		if (this.file == null || !file.canRead()) {
			throw new RuntimeException("File is null or doesn't have read permission");
		}

		this.fileType = getFileType(this.file);
	}

	public GenericExcelReader(File inFile) {
		super();

		if(inFile == null)
			throw new IllegalArgumentException("Invalid file : null");
		
		if (!inFile.canRead()) {
			throw new RuntimeException("file doesn't have read permission");
		}

		this.file = inFile;
		this.fileType = getFileType(this.file);
	}

	/* --- Methods --- */

	public List<T> read(Class<T> entityType) {
		List<T> entityObjects = new ArrayList<T>();

		// Get the workbook instance
		Workbook workbook = getWorkBook();
		int noOfSheets = workbook.getNumberOfSheets();
		logger.debug("Total no of Sheets found : " + noOfSheets);

		// Iterate over all Sheets
		for (int i = 0; i < noOfSheets; i++) {
			entityObjects.addAll(this.read(i, entityType));
		}

		return entityObjects;
	}

	public List<T> read(int sheetNumber, Class<T> entityType) {
		
		try {
			sheetContentsHandler = new ExcelSheetContentsHandler<T>(entityType);

			InputStream inputStream = new FileInputStream(this.file);
			xlsxPackage = OPCPackage.open(inputStream);

			ReadOnlySharedStringsTable sharedStringsTable = new ReadOnlySharedStringsTable(this.xlsxPackage);
			XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);
			StylesTable styles = xssfReader.getStylesTable();
			
			ContentHandler handler = new XSSFSheetXMLHandler(styles, sharedStringsTable, sheetContentsHandler, true);
			
			XSSFReader.SheetIterator worksheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();

			SAXParserFactory saxFactory;
			XMLReader sheetParser;
			for (int i = 0; worksheets.hasNext(); i++) {
				InputStream sheetInpStream = worksheets.next();

				if (i == sheetNumber) {
					saxFactory = SAXParserFactory.newInstance();
					sheetParser = saxFactory.newSAXParser().getXMLReader();
					sheetParser.setContentHandler(handler);
					sheetParser.parse(new InputSource(sheetInpStream));
				}

				IOUtils.closeQuietly(sheetInpStream);
			}

			return sheetContentsHandler.getRowsAsObjects();

		} catch (IOException e) {
			logger.error(e.getMessage());
		} catch (SAXException e) {
			logger.error(e.getMessage());
		} catch (OpenXML4JException e) {
			logger.error(e.getMessage());
		} catch (ParserConfigurationException e) {
			logger.error(e.getMessage());
		}

		return null;
	}

	/* --- Private Helper Methods --- */

	private Workbook getWorkBook() {
		try {
			if (this.fileType == null || this.fileType.isEmpty())
				return null;

			FileInputStream fileStream = new FileInputStream(this.file);

			if (this.fileType.equals(EXCEL_TYPE_XLS)) {
				return new HSSFWorkbook(fileStream);
			} else if (this.fileType.equals(EXCEL_TYPE_XLSX)) {
				return new XSSFWorkbook(fileStream);
			}

			fileStream.close();

		} catch (IOException e) {
			e.printStackTrace();
		}

		throw new RuntimeException("Invalid fileType");
	}

	/**
	 * TODO : Move as a Utility
	 * 
	 * @param inFile
	 * @return
	 */
	private String getFileType(File inFile) {
		// Assuming inFile is not null.. :) as its private method

		String fileName = inFile.getAbsolutePath();

		String extension = "";

		int i = fileName.lastIndexOf('.');
		int p = Math.max(fileName.lastIndexOf('/'), fileName.lastIndexOf('\\'));

		if (i > p) {
			extension = fileName.substring(i + 1);
		}

		return extension;
	}
	
	/* --- Static Utilities --- */
	
	public static Map<String, String> getColumnToPropertyMap(Class<? extends IExcelEntity> entityType) {
		if(entityType == null)
			throw new IllegalArgumentException("Invalid entity type :" + entityType);
		
		Map<String, String> mapping = new HashMap<String, String>();

		Field[] fields = entityType.getDeclaredFields();
		if (fields == null || fields.length == 0) {
			return mapping;
		}
		
		for (int i = 0; i < fields.length; i++) {
			Field f = fields[i];

			Annotation annotation = f.getAnnotation(ExcelColumn.class);
			ExcelColumn ec = (ExcelColumn) annotation;

			if (ec != null && ec.name() != null && !ec.name().isEmpty()) {
				mapping.put(ec.name(), f.getName());
			}
		}

		return Collections.unmodifiableMap(mapping);
	}

}
