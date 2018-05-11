package com.github.millij.eom;

import java.beans.IntrospectionException;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.github.millij.eom.spi.IExcelBean;
import com.github.millij.eom.spi.annotation.ExcelSheet;


public class GenericExcelWriter {

    private static final Logger LOGGER = LoggerFactory.getLogger(GenericExcelWriter.class);

    private final XSSFWorkbook workbook;
    private final OutputStream outputStrem;


    // Constructors
    // ------------------------------------------------------------------------

    public GenericExcelWriter(String filepath) throws FileNotFoundException {
        this(new File(filepath));
    }

    public GenericExcelWriter(File file) throws FileNotFoundException {
        this(new FileOutputStream(file));
    }

    public GenericExcelWriter(OutputStream outputStream) {
        super();

        this.workbook = new XSSFWorkbook();
        this.outputStrem = outputStream;
    }


    // Methods
    // ------------------------------------------------------------------------

    
    // Sheet :: Add
    
    public <EB extends IExcelBean> void addSheet(Class<EB> beanType, List<EB> rowObjects) {
        // Sheet Headers
        List<String> headers = ExcelUtil.getColumnHeaders(beanType);

        this.addSheet(beanType, rowObjects, headers);
    }

    public <EB extends IExcelBean> void addSheet(Class<EB> beanType, List<EB> rowObjects, List<String> headers) {
        // SheetName
        ExcelSheet excelSheet = beanType.getAnnotation(ExcelSheet.class);
        String sheetName = excelSheet.value();

        this.addSheet(beanType, rowObjects, headers, sheetName);
    }

    public <EB extends IExcelBean> void addSheet(Class<EB> beanType, List<EB> rowObjects, String sheetName) {
        // Sheet Headers
        List<String> headers = ExcelUtil.getColumnHeaders(beanType);

        this.addSheet(beanType, rowObjects, headers, sheetName);
    }

    public <EB extends IExcelBean> void addSheet(Class<EB> beanType, List<EB> rowObjects, List<String> headers,
            String sheetName) {
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

            LOGGER.debug("Adding new Sheet[name] to the workbook : {}", sheetName);
            XSSFSheet sheet = workbook.createSheet(sheetName);

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

    private <EB extends IExcelBean> Map<String, List<String>> prepareSheetRowsData(List<String> headers,
            List<EB> rowObjects) throws IllegalAccessException, InvocationTargetException, IntrospectionException {

        final Map<String, List<String>> sheetData = new LinkedHashMap<String, List<String>>();

        // Iterate over Objects
        for (EB excelBean : rowObjects) {
            Map<String, String> row = ExcelUtil.asRowDataMap(excelBean, headers);

            for (String header : headers) {
                List<String> data = sheetData.containsKey(header) ? sheetData.get(header) : new ArrayList<String>();
                String value = row.get(header) != null ? row.get(header) : "";
                data.add(value);

                sheetData.put(header, data);
            }
        }

        return sheetData;
    }



}
