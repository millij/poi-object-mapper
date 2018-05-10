package com.github.millij.eom;

import static com.github.millij.eom.ExcelUtil.EXTN_XLS;
import static com.github.millij.eom.ExcelUtil.EXTN_XLSX;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import com.github.millij.eom.spi.IExcelEntity;
import com.github.millij.eom.spi.handler.ExcelSheetContentsHandler;


public class GenericExcelReader {

    private static final Logger LOGGER = LoggerFactory.getLogger(GenericExcelReader.class);

    private final File file;
    private final String fileType;


    // Constructors
    // ------------------------------------------------------------------------

    public GenericExcelReader(String filePath) {
        super();

        if (StringUtils.isEmpty(filePath)) {
            throw new IllegalArgumentException("GenericExcelReader :: filepath is either empty or null");
        }

        // Set the file
        this.file = new File(filePath);
        if (this.file == null || !file.canRead()) {
            throw new RuntimeException("GenericExcelReader :: File is null or doesn't have read permission");
        }

        this.fileType = ExcelUtil.getFileExtension(this.file);
    }

    public GenericExcelReader(File inFile) {
        super();

        if (inFile == null) {
            throw new IllegalArgumentException("Invalid file : null");
        }

        if (!inFile.canRead()) {
            throw new RuntimeException("GenericExcelReader :: File doesn't have read permission");
        }

        this.file = inFile;
        this.fileType = ExcelUtil.getFileExtension(this.file);
    }



    // Methods
    // ------------------------------------------------------------------------

    public <T extends IExcelEntity> List<T> read(Class<T> entityType) {
        // Get the workbook instance
        Workbook workbook = this.getWorkBook(this.file, this.fileType);
        int noOfSheets = workbook.getNumberOfSheets();
        LOGGER.debug("Total no of Sheets found : " + noOfSheets);

        // Iterate over all Sheets
        List<T> entityObjects = new ArrayList<T>();
        for (int i = 0; i < noOfSheets; i++) {
            entityObjects.addAll(this.read(i, entityType));
        }

        return entityObjects;
    }

    public <T extends IExcelEntity> List<T> read(int sheetNo, Class<T> entityClz) {
        try {
            final OPCPackage opcPkg = OPCPackage.open(this.file);

            // XSSF Reader
            XSSFReader xssfReader = new XSSFReader(opcPkg);
            StylesTable styles = xssfReader.getStylesTable();

            // Content Handler
            ReadOnlySharedStringsTable ssTable = new ReadOnlySharedStringsTable(opcPkg);
            ExcelSheetContentsHandler<T> sheetContentsHandler = new ExcelSheetContentsHandler<T>(entityClz);
            ContentHandler handler = new XSSFSheetXMLHandler(styles, ssTable, sheetContentsHandler, true);

            // XML Reader
            XMLReader xmlReader = SAXHelper.newXMLReader();
            xmlReader.setContentHandler(handler);

            // Iterate over sheets
            XSSFReader.SheetIterator worksheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            for (int i = 0; worksheets.hasNext(); i++) {
                InputStream sheetInpStream = worksheets.next();

                // Read Sheet
                if (i == sheetNo) {
                    xmlReader.parse(new InputSource(sheetInpStream));
                }

                sheetInpStream.close();
            }

            return sheetContentsHandler.getRowsAsObjects();
        } catch (Exception ex) {
            String errMsg =
                    String.format("Error reading sheet %d, entity %s : %s", sheetNo, entityClz, ex.getMessage());
            LOGGER.error(errMsg, ex);
        }

        return null;
    }


    // Private Methods
    // ------------------------------------------------------------------------

    private Workbook getWorkBook(File file, String fileExtn) {
        try {
            FileInputStream fileStream = new FileInputStream(file);
            if (fileExtn.equals(EXTN_XLS)) {
                return new HSSFWorkbook(fileStream);
            } else if (fileExtn.equals(EXTN_XLSX)) {
                return new XSSFWorkbook(fileStream);
            }

            IOUtils.closeQuietly(fileStream);
        } catch (Exception ex) {
            String errMsg = String.format("Failed to Get WorkBook from file - %s : %s", file.toPath(), ex.getMessage());
            LOGGER.error(errMsg, ex);
        }

        throw new RuntimeException("GenericExcelReader :: getWorkBook - Invalid file or fileType passed");
    }


}
