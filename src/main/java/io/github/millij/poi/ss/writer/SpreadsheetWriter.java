package io.github.millij.poi.ss.writer;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

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
import io.github.millij.poi.ss.model.annotations.SheetColumn;
import io.github.millij.poi.util.Beans;
import io.github.millij.poi.util.Spreadsheet;


@Deprecated
public class SpreadsheetWriter {

    private static final Logger LOGGER = LoggerFactory.getLogger(SpreadsheetWriter.class);
    // private static final String DEFAULT_DATE_FORMAT = "dd/MM/yyyy";

    // Default Formats
    private static final String DATE_DEFAULT_FORMAT = "EEE MMM dd HH:mm:ss zzz yyyy";
    private static final String LOCAL_DATE_DEFAULT_FORMAT = "yyyy-MM-dd";

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
        List<String> headers = Spreadsheet.getColumnNames(beanType);

        this.addSheet(beanType, rowObjects, headers);
    }

    public <EB> void addSheet(Class<EB> beanType, List<EB> rowObjects, List<String> headers) {
        // SheetName
        Sheet sheet = beanType.getAnnotation(Sheet.class);
        String sheetName = sheet != null ? sheet.value() : null;

        this.addSheet(beanType, rowObjects, headers, sheetName);
    }

    public <EB> void addSheet(Class<EB> beanType, List<EB> rowObjects, String sheetName) {
        // Sheet Headers
        List<String> headers = Spreadsheet.getColumnNames(beanType);

        this.addSheet(beanType, rowObjects, headers, sheetName);
    }

    public <EB> void addSheet(Class<EB> beanType, List<EB> rowObjects, List<String> headers, String sheetName) {
        // Sanity checks
        if (beanType == null) {
            throw new IllegalArgumentException("GenericExcelWriter :: ExcelBean type should not be null");
        }

        if (CollectionUtils.isEmpty(rowObjects)) {
            LOGGER.error("Skipping excel sheet writing as the ExcelBean collection is empty");
            return;
        }

        if (CollectionUtils.isEmpty(headers)) {
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
            for (int i = 0; i < headers.size(); i++) {
                XSSFCell cell = headerRow.createCell(i);
                cell.setCellValue(headers.get(i));
            }

            // Data Rows
            final Map<String, List<String>> rowsData = this.prepareSheetRowsData(headers, rowObjects);

            final Map<String, String> dateFormatsMap = this.getFormats(beanType);

            for (int i = 0, rowNum = 1; i < rowObjects.size(); i++, rowNum++) {
                final XSSFRow row = sheet.createRow(rowNum);

                int cellNo = 0;
                for (String key : rowsData.keySet()) {
                    final Cell cell = row.createCell(cellNo);

                    final String keyFormat = dateFormatsMap.get(key);
                    if (keyFormat != null) {

                        final String value = rowsData.get(key).get(i);
                        Date date;

                        try {
                            // Date Check
                            final SimpleDateFormat formatter = new SimpleDateFormat(DATE_DEFAULT_FORMAT);
                            date = formatter.parse(value);
                        } catch (ParseException e) {

                            try {
                                // Local Check
                                final SimpleDateFormat formatter = new SimpleDateFormat(LOCAL_DATE_DEFAULT_FORMAT);
                                date = formatter.parse(value);
                            } catch (ParseException ex) {
                                cell.setCellValue(value);
                                cellNo++;
                                continue;
                            }
                        }

                        if (Objects.isNull(date)) {
                            continue;
                        }

                        final SimpleDateFormat formatter = new SimpleDateFormat(keyFormat);
                        final String formattedDate = formatter.format(date);

                        cell.setCellValue(formattedDate);
                        cellNo++;
                        continue;
                    }
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

    private <EB> Map<String, List<String>> prepareSheetRowsData(List<String> headers, List<EB> rowObjects)
            throws Exception {

        final Map<String, List<String>> sheetData = new LinkedHashMap<String, List<String>>();

        // Iterate over Objects
        for (EB excelBean : rowObjects) {
            Map<String, String> row = Spreadsheet.asRowDataMap(excelBean, headers);

            for (String header : headers) {
                List<String> data = sheetData.containsKey(header) ? sheetData.get(header) : new ArrayList<String>();
                String value = row.get(header) != null ? row.get(header) : "";
                data.add(value);

                sheetData.put(header, data);
            }
        }

        return sheetData;
    }

    public static Map<String, String> getFormats(Class<?> beanType) {

        if (beanType == null) {
            throw new IllegalArgumentException("getColumnToPropertyMap :: Invalid ExcelBean type - " + beanType);
        }

        final Map<String, String> headFormatMap = new HashMap<String, String>();

        // Fields
        final Field[] fields = beanType.getDeclaredFields();

        for (Field f : fields) {
            if (!f.isAnnotationPresent(SheetColumn.class)) {
                continue;
            }

            final SheetColumn ec = f.getAnnotation(SheetColumn.class);

            if (ec != null) {
                final Class<?> fieldType = f.getType();
                if (fieldType == Date.class || fieldType == LocalDate.class) {
                    final String value = ec.value();
                    headFormatMap.put(StringUtils.isNotBlank(value) ? value : f.getName(), ec.format());
                }
            }
        }

        // Methods
        final Method[] methods = beanType.getDeclaredMethods();

        for (Method m : methods) {
            if (!m.isAnnotationPresent(SheetColumn.class)) {
                continue;
            }

            final String fieldName = Beans.getFieldName(m);
            final Class<?> objType = m.getReturnType();

            final SheetColumn ec = m.getAnnotation(SheetColumn.class);
            if (objType == Date.class || objType == LocalDate.class) {
                final String value = StringUtils.isBlank(ec.value()) ? fieldName : ec.value();
                headFormatMap.put(value, ec.format());
            }
        }
        return headFormatMap;
    }
}
