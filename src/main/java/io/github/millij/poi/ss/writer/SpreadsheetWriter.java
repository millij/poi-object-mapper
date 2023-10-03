package io.github.millij.poi.ss.writer;

import io.github.millij.poi.ss.model.annotations.Sheet;
import io.github.millij.poi.ss.model.annotations.SheetColumn;
import io.github.millij.poi.util.Spreadsheet;
import io.github.millij.poi.ss.reader.SpreadsheetReader;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


@Deprecated
public class SpreadsheetWriter {

    private static final Logger LOGGER = LoggerFactory.getLogger(SpreadsheetWriter.class);
    // private static final String DEFAULT_DATE_FORMAT = "dd/MM/yyyy";

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
            Map<String, List<String>> rowsData = this.prepareSheetRowsData(headers, rowObjects);
            for (int i = 0, rowNum = 1; i < rowObjects.size(); i++, rowNum++) {
                final XSSFRow row = sheet.createRow(rowNum);

                final Map<String, String> dateFormatsMap = this.getFormats(beanType);

                final List<String> formulaCols = this.getFormulaCols(beanType);

                int cellNo = 0;
                for (String key : rowsData.keySet()) {
                    Cell cell = row.createCell(cellNo);

                    String keyFormat = dateFormatsMap.get(key);
                    if (keyFormat != null) {

                        try {

                            String value = rowsData.get(key).get(i);

                            LocalDate localDate = LocalDate.parse(value,
                                    DateTimeFormatter.ofPattern(SpreadsheetReader.DEFAULT_DATE_FORMAT));

                            DateTimeFormatter formatter = DateTimeFormatter.ofPattern(keyFormat);

                            String formattedDateTime = localDate.format(formatter);

                            cell.setCellValue(formattedDateTime);
                            cellNo++;

                        } catch (Exception e) {
                            e.printStackTrace();
                        }

                    } else if (formulaCols.contains(key)) {
                        String value = rowsData.get(key).get(i);
                        cell.setCellFormula(value);
                        cellNo++;

                    } else {
                        String value = rowsData.get(key).get(i);
                        cell.setCellValue(value);
                        cellNo++;
                    }
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

        Map<String, String> headFormatMap = new HashMap<String, String>();

        // Fields
        final Field[] fields = beanType.getDeclaredFields();

        for (Field f : fields) {

            SheetColumn ec = f.getAnnotation(SheetColumn.class);

            if (ec != null && StringUtils.isNotEmpty(ec.value())) {
                if (ec.isFormatted())
                    headFormatMap.put(ec.value(), ec.format());
            }
        }

        // Methods
        final Method[] methods = beanType.getDeclaredMethods();

        for (Method m : methods) {

            SheetColumn ec = m.getAnnotation(SheetColumn.class);

            if (ec != null && StringUtils.isNotEmpty(ec.value())) {
                if (ec.isFormatted())
                    headFormatMap.put(ec.value(), ec.format());
            }
        }

        return headFormatMap;
    }


    public static List<String> getFormulaCols(Class<?> beanType) {
        if (beanType == null) {
            throw new IllegalArgumentException("getColumnToPropertyMap :: Invalid ExcelBean type - " + beanType);
        }

        List<String> formulaCols = new ArrayList<String>();

        // Fields
        final Field[] fields = beanType.getDeclaredFields();

        for (Field f : fields) {

            SheetColumn ec = f.getAnnotation(SheetColumn.class);

            if (ec != null && StringUtils.isNotEmpty(ec.value())) {
                if (ec.isFormula())
                    formulaCols.add(ec.value());
            }
        }

        // Methods
        final Method[] methods = beanType.getDeclaredMethods();

        for (Method m : methods) {

            SheetColumn ec = m.getAnnotation(SheetColumn.class);

            if (ec != null && StringUtils.isNotEmpty(ec.value())) {
                if (ec.isFormula())
                    formulaCols.add(ec.value());
            }
        }

        return formulaCols;

    }



}
