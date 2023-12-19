package io.github.millij.poi.util;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import io.github.millij.poi.ss.model.Column;
import io.github.millij.poi.ss.model.DateTimeType;
import io.github.millij.poi.ss.model.annotations.Sheet;
import io.github.millij.poi.ss.model.annotations.SheetColumn;


/**
 * Spreadsheet related utilities.
 */
public final class Spreadsheet {

    private static final Logger LOGGER = LoggerFactory.getLogger(Spreadsheet.class);

    private Spreadsheet() {
        super();
        // Utility Class
    }


    //
    // Cell
    // ------------------------------------------------------------------------

    /**
     * Splits the CellReference and returns only the column reference.
     * 
     * @param cellRef the cell reference value (ex. D3)
     * 
     * @return returns the column index "D" from the cell reference "D3"
     */
    public static String getCellColumnReference(final String cellRef) {
        final String cellColRef = cellRef.split("[0-9]*$")[0];
        return cellColRef;
    }


    //
    // Sheet & SheetColumn Annotations
    // ------------------------------------------------------------------------

    /**
     * Get the name of the sheet as defined in the {@link SheetColumn} annotation.
     * 
     * @param beanType The bean Type
     * 
     * @return Sheet name as String.
     */
    public static String getSheetName(final Class<?> beanType) {
        final Sheet sheet = beanType.getAnnotation(Sheet.class);
        final String sheetName = Objects.isNull(sheet) ? null : sheet.value();
        return sheetName;
    }

    public static String getSheetColumnName(final SheetColumn sheetColumn, final String defaultName) {
        // Name
        final String scValue = sheetColumn.value();
        final String colName = Objects.isNull(scValue) || scValue.isBlank() ? defaultName : scValue;

        return colName;
    }

    /**
     * Prepare {@link Column} def from {@link SheetColumn} annotation info.
     * 
     * @param sheetCol instance of the {@link SheetColumn} annotation
     * @param defaultName default name of the Column
     * 
     * @return {@link Column} object representation.
     */
    public static Column asColumn(final SheetColumn sheetCol, final String defaultName) {
        // Sanity checks
        if (Objects.isNull(sheetCol)) {
            return new Column(defaultName);
        }

        // Column Name
        final String colName = Spreadsheet.getSheetColumnName(sheetCol, defaultName);

        // Prepare Column
        final Column column = new Column();
        column.setName(colName);
        column.setNullable(sheetCol.nullable());
        column.setFormat(sheetCol.format());
        column.setOrder(sheetCol.order());
        column.setDatetimeType(sheetCol.datetime());

        return column;
    }


    //
    // Bean :: Property Utils
    // ------------------------------------------------------------------------

    public static Map<String, Column> getPropertyToColumnDefMap(final Class<?> beanType) {
        // Sanity checks
        if (Objects.isNull(beanType)) {
            final String errMsg = String.format("#getPropertyToColumnDefMap :: Input type is NULL");
            throw new IllegalArgumentException(errMsg);
        }

        // Property to Column name Mapping
        final Map<String, Column> mappings = new HashMap<>();

        // Fields
        final Field[] fields = beanType.getDeclaredFields();
        for (final Field f : fields) {
            final SheetColumn sc = f.getAnnotation(SheetColumn.class);
            if (Objects.isNull(sc)) {
                continue; // Skip it here as the annotation may be present on getter/setter
            }

            // Field Name
            final String fieldName = f.getName();

            // Column
            final Column column = Spreadsheet.asColumn(sc, fieldName);
            mappings.put(fieldName, column);
        }

        // Methods
        final Method[] methods = beanType.getDeclaredMethods();
        for (final Method m : methods) {
            final String fieldName = Beans.getFieldName(m);
            if (mappings.containsKey(fieldName)) {
                continue; // Skip it as it already exists from Field defs
            }

            // Annotation
            final SheetColumn sc = m.getAnnotation(SheetColumn.class);
            if (Objects.isNull(sc) && m.getName().startsWith("set")) {
                continue; // Skip setter
            }

            // Column
            final Column column = Spreadsheet.asColumn(sc, fieldName);
            mappings.put(fieldName, column);
        }

        return Collections.unmodifiableMap(mappings);
    }


    // Read from Bean : as Row Data
    // ------------------------------------------------------------------------

    public static Map<String, String> asRowDataMap(final Object beanObj, final List<String> colHeaders)
            throws Exception {
        // Excel Bean Type
        final Class<?> beanType = beanObj.getClass();

        // RowData map
        final Map<String, String> rowDataMap = new HashMap<String, String>();

        // Fields
        for (final Field f : beanType.getDeclaredFields()) {
            if (!f.isAnnotationPresent(SheetColumn.class)) {
                continue;
            }

            final String fieldName = f.getName();
            final SheetColumn sc = f.getAnnotation(SheetColumn.class);

            final String header = Spreadsheet.getSheetColumnName(sc, fieldName);
            if (!colHeaders.contains(header)) {
                continue;
            }

            rowDataMap.put(header, Beans.getFieldValueAsString(beanObj, fieldName));
        }

        // Methods
        for (final Method m : beanType.getDeclaredMethods()) {
            if (!m.isAnnotationPresent(SheetColumn.class)) {
                continue;
            }

            final String fieldName = Beans.getFieldName(m);
            final SheetColumn sc = m.getAnnotation(SheetColumn.class);

            final String header = Spreadsheet.getSheetColumnName(sc, fieldName);
            if (!colHeaders.contains(header)) {
                continue;
            }

            rowDataMap.put(header, Beans.getFieldValueAsString(beanObj, fieldName));
        }

        return rowDataMap;
    }


    // Write to Bean :: from Row data
    // ------------------------------------------------------------------------

    public static <T> T rowAsBean(Class<T> beanClz, Map<String, Column> propColumnMap,
            Map<String, String> headerCellRefsMap, Map<String, Object> rowDataMap) {
        // Sanity checks
        if (Objects.isNull(headerCellRefsMap) || Objects.isNull(rowDataMap)) {
            return null;
        }

        // Validate
        final boolean isValidRowData = Spreadsheet.validateRowData(rowDataMap, headerCellRefsMap, propColumnMap);
        if (!isValidRowData) {
            LOGGER.debug("#rowAsBean :: Skipping the bean creation as the ROW data in INVALID");
            return null;
        }

        try {
            // Create new Instance
            final T bean = beanClz.getDeclaredConstructor().newInstance();

            for (final String propName : propColumnMap.keySet()) {
                // Prop Column Definition
                final Column propColDef = propColumnMap.get(propName);
                final String propColName = propColDef.getName();

                // Get the Header Cell Ref
                final String normalizedColName = Spreadsheet.normalize(propColName);
                final String propCellRef = headerCellRefsMap.get(normalizedColName);
                if (Objects.isNull(propCellRef) || propCellRef.isBlank()) {
                    LOGGER.debug("{} :: No Cell Ref found [Prop - Col] : [{} - {}]", beanClz, propName, propColName);
                    continue;
                }

                // Property Value and Format
                final Object propValue = rowDataMap.get(propCellRef);
                final String dataFormat = propColDef.getFormat();
                final DateTimeType datetimeType = propColDef.getDatetimeType();

                // Set Value
                try {
                    // Set the property value in the current row object bean
                    Beans.setProperty(bean, propName, propValue, dataFormat, datetimeType);
                } catch (Exception ex) {
                    String errMsg = String.format("Failed to set bean property - %s, value - %s", propName, propValue);
                    LOGGER.error(errMsg, ex);
                }

            }

            return bean;
        } catch (Exception ex) {
            String errMsg = String.format("Error while creating bean - %s, from - %s", beanClz, rowDataMap);
            LOGGER.error(errMsg, ex);
        }

        return null;
    }

    private static boolean validateRowData(final Map<String, Object> rowDataMap,
            final Map<String, String> headerCellRefsMap, final Map<String, Column> propColumnMap) {
        // Good Values
        int noOfValuesFound = 0;

        //
        for (final String propName : propColumnMap.keySet()) {
            // Prop Column Definition
            final Column propColDef = propColumnMap.get(propName);
            final String propColName = propColDef.getName();
            final String normalizedColName = Spreadsheet.normalize(propColName);

            // Get the Header Cell Ref
            final String propCellRef = headerCellRefsMap.containsKey(propColName) //
                    ? headerCellRefsMap.get(propColName) //
                    : headerCellRefsMap.get(normalizedColName);
            if (Objects.isNull(propCellRef) || propCellRef.isBlank()) {
                continue;
            }

            // Property Value and Format
            final Object propValue = rowDataMap.get(propCellRef);
            if (Objects.isNull(propValue)) {
                continue;
            }

            // TODO :: Handle FORMULAs

            noOfValuesFound++;
        }

        final boolean hasAnyValuePresent = noOfValuesFound > 0;
        return hasAnyValuePresent;
    }


    // Other Methods
    // ------------------------------------------------------------------------

    /**
     * Normalize the string. typically used for case-insensitive comparison.
     */
    public static String normalize(final String inStr) {
        // Sanity checks
        if (Objects.isNull(inStr)) {
            return "";
        }

        // Special characters
        final String cleanStr = inStr.replaceAll("–", " ").replaceAll("[-\\[\\]/{}:.,;#%=()*+?\\^$|<>&\"\'\\\\]", " ");
        final String normalizedStr = cleanStr.toLowerCase().trim().replaceAll("\\s+", "_");

        return normalizedStr;
    }


}
