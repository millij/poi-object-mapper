package io.github.millij.poi.util;

import io.github.millij.poi.SpreadsheetReadException;
import io.github.millij.poi.ss.model.annotations.SheetColumn;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * Spreadsheet related utilities.
 */
public final class Spreadsheet {

	/**
	 * Splits the CellReference and returns only the column reference.
	 * 
	 * @param cellRef the cell reference value (ex. D3)
	 * @return returns the column index "D" from the cell reference "D3"
	 */

	private static final Logger LOGGER = LoggerFactory.getLogger(Spreadsheet.class);

	private Spreadsheet() {
		// Utility Class
	}

	// Utilities
	// ------------------------------------------------------------------------

	/**
	 * Splits the CellReference and returns only the column reference.
	 * 
	 * @param cellRef the cell reference value (ex. D3)
	 * @return returns the column index "D" from the cell reference "D3"
	 */
	public static String getCellColumnReference(String cellRef) {
		String cellColRef = cellRef.split("[0-9]*$")[0];
		return cellColRef;
	}

	// Bean :: Property Utils

	public static Map<String, String> getPropertyToColumnNameMap(Class<?> beanType) {
		// Sanity checks
		if (beanType == null) {
			throw new IllegalArgumentException("getColumnToPropertyMap :: Invalid ExcelBean type - " + beanType);
		}

		// Property to Column name Mapping
		final Map<String, String> mapping = new HashMap<String, String>();

		// Fields
		Field[] fields = beanType.getDeclaredFields();
		for (Field f : fields) {
			String fieldName = f.getName();
			mapping.put(fieldName, fieldName);

			SheetColumn ec = f.getAnnotation(SheetColumn.class);
			if (ec != null && StringUtils.isNotEmpty(ec.value())) {
				mapping.put(fieldName, ec.value());
			}
		}

		// Methods
		Method[] methods = beanType.getDeclaredMethods();
		for (Method m : methods) {
			String fieldName = Beans.getFieldName(m);
			if (!mapping.containsKey(fieldName)) {
				mapping.put(fieldName, fieldName);
			}

			SheetColumn ec = m.getAnnotation(SheetColumn.class);
			if (ec != null && StringUtils.isNotEmpty(ec.value())) {
				mapping.put(fieldName, ec.value());
			}
		}

		LOGGER.info("Bean property to Excel Column of - {} : {}", beanType, mapping);
		return Collections.unmodifiableMap(mapping);
	}

	public static Map<String, String> getColumnToPropertyMap(Class<?> beanType) {
		// Column to Property Mapping
		final Map<String, String> columnToPropMap = new HashMap<String, String>();

		// Bean Property to Column Mapping
		final Map<String, String> propToColumnMap = getPropertyToColumnNameMap(beanType);
		for (String prop : propToColumnMap.keySet()) {
			columnToPropMap.put(propToColumnMap.get(prop), prop);
		}

		LOGGER.info("Excel Column to property map of - {} : {}", beanType, columnToPropMap);
		return Collections.unmodifiableMap(columnToPropMap);
	}

	public static List<String> getColumnNames(Class<?> beanType) {
		// Bean Property to Column Mapping
		final Map<String, String> propToColumnMap = getPropertyToColumnNameMap(beanType);
		final Map<Integer, String> indexToPropMap = getIndexToPropertyMap(beanType);

		Set<Integer> indexes = indexToPropMap.keySet();
		List<Integer> indexList = new ArrayList<Integer>(indexes);
		Collections.sort(indexList);

		ArrayList<String> columnNames = new ArrayList<>();

		for (Integer index : indexList) {
			String colValue = propToColumnMap.get(indexToPropMap.get(index));
			columnNames.add(colValue);
		}
		return columnNames;
	}

	// Read from Bean : as Row Data
	// ------------------------------------------------------------------------

	public static Map<String, String> asRowDataMap(Object beanObj, List<String> colHeaders) throws Exception {
		// Excel Bean Type
		final Class<?> beanType = beanObj.getClass();

		// RowData map
		final Map<String, String> rowDataMap = new HashMap<String, String>();

		// Fields
		for (Field f : beanType.getDeclaredFields()) {
			if (!f.isAnnotationPresent(SheetColumn.class)) {
				continue;
			}

			String fieldName = f.getName();

			SheetColumn ec = f.getAnnotation(SheetColumn.class);
			String header = StringUtils.isEmpty(ec.value()) ? fieldName : ec.value();
			if (!colHeaders.contains(header)) {
				continue;
			}

			rowDataMap.put(header, Beans.getFieldValueAsString(beanObj, fieldName));
		}

		// Methods
		for (Method m : beanType.getDeclaredMethods()) {
			if (!m.isAnnotationPresent(SheetColumn.class)) {
				continue;
			}

			String fieldName = Beans.getFieldName(m);

			SheetColumn ec = m.getAnnotation(SheetColumn.class);
			String header = StringUtils.isEmpty(ec.value()) ? fieldName : ec.value();
			if (!colHeaders.contains(header)) {
				continue;
			}

			rowDataMap.put(header, Beans.getFieldValueAsString(beanObj, fieldName));
		}

		return rowDataMap;
	}

	// Write to Bean :: from Row data
	// ------------------------------------------------------------------------

	public static <T> T rowAsBean(Class<T> beanClz, Map<String, String> cellProperies, Map<String, Object> cellValues) {
		// Sanity checks
		if (cellValues == null || cellProperies == null) {
			return null;
		}

		try {
			// Create new Instance
			T rowBean = beanClz.newInstance();

			// Fill in the data
			for (String cellName : cellProperies.keySet()) {
				String propName = cellProperies.get(cellName);
				if (StringUtils.isEmpty(propName)) {
					LOGGER.debug("{} : No mathching property found for column[name] - {} ", beanClz, cellName);
					continue;
				}

				Object propValue = cellValues.get(cellName);
				try {
					// Set the property value in the current row object bean
					BeanUtils.setProperty(rowBean, propName, propValue);
				} catch (IllegalAccessException | InvocationTargetException ex) {
					String errMsg = String.format("Failed to set bean property - %s, value - %s", propName, propValue);
					LOGGER.error(errMsg, ex);
				}
			}

			return rowBean;
		} catch (Exception ex) {
			String errMsg = String.format("Error while creating bean - %s, from - %s", beanClz, cellValues);
			LOGGER.error(errMsg, ex);
		}

		return null;
	}

	public static Map<Integer, String> getIndexToPropertyMap(Class<?> beanType) {
		// Sanity checks
		if (Objects.isNull(beanType)) {
			throw new IllegalArgumentException("getColumnIndexToPropertyMap :: Invalid ExcelBean type - " + beanType);
		}

		// Column Index to Property Mapping
		final Map<Integer, String> mapping = new HashMap<Integer, String>();

		// Fields
		Field[] fields = beanType.getDeclaredFields();
		for (Field f : fields) {
			String fieldName = f.getName();

			SheetColumn ec = f.getAnnotation(SheetColumn.class);
			try {
				if (ec != null && ec.index() == -1) {
					throw new NullPointerException(
							"Index must be intialized at annotation level for field " + fieldName);
				} else {
					mapping.put(ec.index(), fieldName);
				}
			} catch (NullPointerException ex) {
				if (!Objects.isNull(ex.getMessage())) {
					LOGGER.error(ex.getMessage());
				}
			}
		}

		// Methods
		Method[] methods = beanType.getDeclaredMethods();
		for (Method m : methods) {
			String fieldName = Beans.getFieldName(m);

			SheetColumn ec = m.getAnnotation(SheetColumn.class);

			try {
				if (ec != null && ec.index() == -1) {
					throw new NullPointerException(
							"Index must be intialized at annotation level for field " + fieldName);
				} else {
					mapping.put(ec.index(), fieldName);
				}
			} catch (NullPointerException ex) {
				if (!Objects.isNull(ex.getMessage())) {
					LOGGER.error(ex.getMessage());
				}
			}
		}

		LOGGER.info("Java Bean Index to Bean Property - {} : {}", beanType, mapping);
		return mapping;
	}

	// Write to Bean :: from Row data
	// ------------------------------------------------------------------------

	public static <T> T rowAsBean(Class<T> beanClz, Map<String, Object> fieldRowValues) {
		// Sanity checks
		if (fieldRowValues == null || beanClz == null) {
			return null;
		}

		try {
			// Create new Instance
			T rowBean = beanClz.newInstance();

			// Fill in the data
			for (String fieldName : fieldRowValues.keySet()) {
				Object propValue = fieldRowValues.get(fieldName);

				try {
					// Set the property value in the current row object bean
					BeanUtils.setProperty(rowBean, fieldName, propValue);
				} catch (IllegalAccessException | InvocationTargetException ex) {
					String errMsg = String.format("Failed to set bean property - %s, value - %s", fieldName, propValue);
					LOGGER.error(errMsg, ex);
				}
			}

			return rowBean;
		} catch (Exception ex) {
			String errMsg = String.format("Error while creating bean - %s, from - %s", beanClz, fieldRowValues);
			LOGGER.error(errMsg, ex);
		}

		return null;
	}

	public static <T> boolean isIndexed(Class<T> beanClz,Map<Integer, String> headerMap){
		// Sanity checks
		if (Objects.isNull(beanClz)) {
			throw new IllegalArgumentException("isIndexed :: Invalid ExcelBean type - " + beanClz);
		}
		
		Map<Integer,String> indexFieldMap = new HashMap<Integer,String>();
		
		//Fields
		Field[] fields = beanClz.getDeclaredFields();
		for(Field f : fields) {
			String fieldName = f.getName();
			SheetColumn ec = f.getAnnotation(SheetColumn.class);
			if(!Objects.isNull(ec)) {
				indexFieldMap.put(ec.index(),fieldName);
			}
		}
		
		//Methods
		Method[] methods = beanClz.getDeclaredMethods();
		for(Method m : methods) {
			String fieldName = Beans.getFieldName(m);
			SheetColumn ec = m.getAnnotation(SheetColumn.class);
			if(!Objects.isNull(ec)) {
				indexFieldMap.put(ec.index(),fieldName);
			}
		}
		
		if(indexFieldMap.size() == headerMap.size()) {
			return true;			
		}
		return false;

	}

}
