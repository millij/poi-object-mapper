package io.github.millij.poi.ss.writer;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.stream.Collectors;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import io.github.millij.poi.ss.model.annotations.Sheet;
import io.github.millij.poi.util.Spreadsheet;

@Deprecated
public class SpreadsheetWriter {

    private static final Logger LOGGER = LoggerFactory.getLogger(SpreadsheetWriter.class);

    private final XSSFWorkbook workbook;
    private final OutputStream outputStrem;

    
    // Constructors
    // ------------------------------------------------------------------------

    public SpreadsheetWriter(String filepath) throws FileNotFoundException {
	this(new File(filepath));
    }

    public SpreadsheetWriter(File file) throws FileNotFoundException {
	this(new FileOutputStream(file));
    }

    public SpreadsheetWriter(OutputStream outputStream) {
	super();

	this.workbook = new XSSFWorkbook();
	this.outputStrem = outputStream;
    }

    
    // Methods
    // ------------------------------------------------------------------------

    // Sheet :: Add

    public <EB> void addSheet(Class<EB> beanType, List<EB> rowObjects) {
	// Sheet Headers
	Map<String, String> headerMap = Spreadsheet.getPropertyToColumnNameMap(beanType);

	this.addSheet(beanType, rowObjects, headerMap);
    }

    public <EB> void addSheet(Class<EB> beanType, List<EB> rowObjects, Map<String, String> headerMap) {
	// SheetName
	Sheet sheet = beanType.getAnnotation(Sheet.class);
	String sheetName = sheet != null ? sheet.value() : null;

	this.addSheet(beanType, rowObjects, headerMap, sheetName);
    }

    public <EB> void addSheet(Class<EB> beanType, List<EB> rowObjects, Set<String> headers) {
	// SheetName
	Sheet sheet = beanType.getAnnotation(Sheet.class);
	String sheetName = sheet != null ? sheet.value() : null;

	// Sheet Headers
	Map<String, String> headerMap = Spreadsheet.getPropertyToColumnNameMap(beanType);
	Map<String, String> headerMapCopy = headerMap.entrySet().stream()
		.collect(Collectors.toMap(Map.Entry::getKey, Map.Entry::getValue));
	headerMapCopy.values().retainAll(headers);

	this.addSheet(beanType, rowObjects, Collections.unmodifiableMap(headerMapCopy), sheetName);
    }

    public <EB> void addSheet(Class<EB> beanType, List<EB> rowObjects, String sheetName) {
        // Sheet Headers
    	Map<String, String> headerMap = Spreadsheet.getPropertyToColumnNameMap(beanType);

        this.addSheet(beanType, rowObjects, headerMap, sheetName);
    }

    public <EB> void addSheet(Class<EB> beanType, List<EB> rowObjects, Map<String, String> headerMap,
	    String sheetName) {
	// Sanity checks
	if (beanType == null) {
	    throw new IllegalArgumentException("GenericExcelWriter :: ExcelBean type should not be null");
	}

	if (CollectionUtils.isEmpty(rowObjects)) {
	    LOGGER.error("Skipping excel sheet writing as the ExcelBean collection is empty");
	    return;
	}

	if (headerMap == null | headerMap.isEmpty()) {
	    LOGGER.error("Skipping excel sheet writing as the headers collection is empty");
	    return;
	}

	try {
	    XSSFSheet exSheet = workbook.getSheet(sheetName);
	    if (exSheet != null) {
		String errMsg = String.format("A Sheet with the passed name already exists : %s", sheetName);
		throw new IllegalArgumentException(errMsg);
	    }

	    XSSFSheet sheet = StringUtils.isEmpty(sheetName) ? workbook.createSheet() : workbook.createSheet(sheetName);
	    LOGGER.debug("Added new Sheet[name] to the workbook : {}", sheet.getSheetName());

	    // Header
	    XSSFRow headerRow = sheet.createRow(0);
	    Iterator<Entry<String, String>> iterator = headerMap.entrySet().iterator();
	    for (int i = 0; iterator.hasNext(); i++) {
		XSSFCell cell = headerRow.createCell(i);
		cell.setCellValue(headerMap.get(iterator.next().getKey()));
	    }

	    // Data Rows
	    Map<String, List<String>> rowsData = this.prepareSheetRowsData(headerMap, rowObjects);
	    for (int i = 0, rowNum = 1; i < rowObjects.size(); i++, rowNum++) {
		final XSSFRow row = sheet.createRow(rowNum);

		int cellNo = 0;
		for (String key : rowsData.keySet()) {
		    Cell cell = row.createCell(cellNo);
		    String value = rowsData.get(key).get(i);
		    cell.setCellValue(value);
		    cellNo++;
		}
	    }

	} catch (Exception ex) {
	    String errMsg = String.format("Error while preparing sheet with passed row objects : %s", ex.getMessage());
	    LOGGER.error(errMsg, ex);
	}
    }

    // Sheet :: Append to existing

    // Write

    public void write() throws IOException {
	workbook.write(outputStrem);
	workbook.close();
    }

    // Private Methods
    // ------------------------------------------------------------------------

    private <EB> Map<String, List<String>> prepareSheetRowsData(Map<String, String> headerMap, List<EB> rowObjects)
	    throws Exception {

	final Map<String, List<String>> sheetData = new LinkedHashMap<String, List<String>>();

	// Iterate over Objects
	for (EB excelBean : rowObjects) {
	    Map<String, String> row = Spreadsheet.asRowDataMap(excelBean, headerMap);

	    for (String header : headerMap.values()) {
		List<String> data = sheetData.containsKey(header) ? sheetData.get(header) : new ArrayList<String>();
		String value = row.get(header) != null ? row.get(header) : "";
		data.add(value);

		sheetData.put(header, data);
	    }
	}

	return sheetData;
    }

}
