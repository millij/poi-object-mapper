package com.eom.handler;

import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;

import com.eom.GenericExcelReader;
import com.eom.IExcelEntity;

public class ExcelSheetContentsHandler<T extends IExcelEntity> implements SheetContentsHandler {

	private static final Log logger = LogFactory.getLog(ExcelSheetContentsHandler.class);

	private Class<T> entityType;
	private int noOfRowsToSkip = 0;

	private boolean verifyHeader = true;
	private int headerRow = 0;

	private int currentRow = 0;

	private Map<String, String> entityPropertyMapping = new HashMap<String, String>();
	private T currentRowEntity;
	private Map<String, String> headerCellMap = new HashMap<String, String>();
	private List<T> rowObjList = new ArrayList<T>();

	/* --- Constructors --- */

	public ExcelSheetContentsHandler(Class<T> entityType) {
		super();
		
		if(entityType == null)
			throw new IllegalArgumentException("Invalid entity type :" + entityType);
		
		this.entityType = entityType;
		
		// Entity Cell Mapping
		this.entityPropertyMapping = validateAndGetEntityColumnMapping();
	}

	public ExcelSheetContentsHandler(Class<T> entityType, int noOfRowsToSkip) {
		this(entityType);
		this.noOfRowsToSkip = noOfRowsToSkip;
	}
	
	
	public ExcelSheetContentsHandler(Class<T> entityType, int noOfRowsToSkip, int headerRow) {
		this(entityType, noOfRowsToSkip);
		this.headerRow = headerRow;
	}
	
	/* --- Methods --- */

	/**
	 * Returns the List of Object that read from the Excel Sheet.
	 * 
	 * @return Objects List
	 */
	public List<T> getRowsAsObjects() {
		if (rowObjList == null)
			rowObjList = new ArrayList<T>();

		return rowObjList;
	}

	public boolean isVerifyHeader() {
		return verifyHeader;
	}

	public void setVerifyHeader(boolean verifyHeader) {
		this.verifyHeader = verifyHeader;
	}

	/* --- Inherited Implementations --- */

	@Override
	public void startRow(int rowNum) {
		this.currentRow = rowNum;
		
		if (rowNum == headerRow || rowNum < noOfRowsToSkip)
			return;
		
		// Create a new instance of entity for each row
		currentRowEntity = this.getEntityNewInstance();
	}

	@Override
	public void cell(String cellReference, String formattedValue) {
		if (this.currentRow < noOfRowsToSkip)
			return;

		if (formattedValue == null || formattedValue.length() == 0)
			return;

		// Handle the Header Row
		if (this.verifyHeader && this.headerRow == this.currentRow) {
			this.saveHeaderCellValue(cellReference, formattedValue);
		} else {
			this.saveRowCellValue(cellReference, formattedValue);
		}
		
	}
	
	@Override
	public void endRow() {
		if (this.headerRow == this.currentRow && verifyHeader)
			this.verifySheetHeader();

		
		if (currentRow < noOfRowsToSkip)
			return;
		
		// Add the current row entity to the Objects list
		if (currentRowEntity != null) {
			this.rowObjList.add(currentRowEntity);
		}

		// reset current Row object to null
		currentRowEntity = null;
	}


	@Override
	public void headerFooter(String text, boolean isHeader, String tagName) {
		// TODO Auto-generated method stub

	}

	/* --- Private Helpers --- */

	private T getEntityNewInstance() {
		try {
			return entityType.newInstance();
		} catch (InstantiationException e) {
			logger.error(e.getMessage());
		} catch (IllegalAccessException e) {
			logger.error(e.getMessage());
		}
		
		return null;
	}
	
	private Map<String, String> validateAndGetEntityColumnMapping() {
		Map<String, String> mapping = GenericExcelReader.getColumnToPropertyMap(entityType);
		if(mapping == null) {
			mapping = new HashMap<String, String>();
		}
		
		return mapping;
	}

	
	/**
	 * TODO : complete
	 */
	private void verifySheetHeader() {
		
		
	}
	
	private void saveRowCellValue(String cellRef, String cellValue) {
		
		// Sanity Checks
		if(this.currentRowEntity == null) {
			logger.warn("How come the row entity is NULL ??? Verify once");
			return;
		}
		
		if (cellRef == null || cellRef.length() == 0) {
			logger.error("Cell reference is null or empty");
			return;
		}
		
		if (cellValue == null || cellValue.length() == 0) {
			logger.warn("Cell's formatted value is null or empty");
			return;
		}
		
		// Now set the value in the row entity
		String cellColRef = getCellColumnReference(cellRef);
		String cellColName = headerCellMap.get(cellColRef);	
		
		if(cellColName == null)
			return;
		
		String entityPropName = this.entityPropertyMapping.get(cellColName);
		if(entityPropName == null || entityPropName.isEmpty()) {
			logger.debug("No mathching property is found / is mapped for column with name :" + cellColName);
			return;
		}
		
		// Set the property value in the current row object bean
		try {
			
			BeanUtils.setProperty(this.currentRowEntity, entityPropName, cellValue);
			
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		} catch (InvocationTargetException e) {
			e.printStackTrace();
		}
	}

	private void saveHeaderCellValue(String cellRef, String cellValue) {
		String cellColRef = getCellColumnReference(cellRef);
		
		// Put the header name and cell reference in the map
		headerCellMap.put(cellColRef, cellValue);
	}
	
	private String getCellColumnReference(String cellRef) {
		if (cellRef == null || cellRef.length() == 0) {
			return "";
		}
		
		// Splits the Cell name and returns the column reference
		String cellColRef = cellRef.split("[0-9]*$")[0];
		
		return cellColRef;
	}

}
