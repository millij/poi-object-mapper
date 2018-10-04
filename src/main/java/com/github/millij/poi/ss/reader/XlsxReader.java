package com.github.millij.poi.ss.reader;

import java.io.File;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import com.github.millij.poi.ExcelReadException;
import com.github.millij.poi.ss.handler.XSSFSheetContentsHandler;


/**
 * Reader impletementation of {@link Workbook} for an OOXML .xlsx file. This implementation is
 * suitable for low memory sax parsing or similar.
 * 
 * @see XlsxStreamReader
 */
public class XlsxReader extends WorkbookReader {

    private static final Logger LOGGER = LoggerFactory.getLogger(XlsxReader.class);


    // Constructor

    public XlsxReader() {
        super(WORKBOOK_XLSX, false);
    }


    // WorkbookReader Impl
    // ------------------------------------------------------------------------

    @Override
    public <EB> List<EB> read(File file, int sheetNo, Class<EB> beanClz) throws ExcelReadException {
        // Sanity checks
        if (beanClz == null) {
            throw new IllegalArgumentException("XlsReader :: Bean class type should not be null");
        }

        try {
            final OPCPackage opcPkg = OPCPackage.open(file);

            // XSSF Reader
            XSSFReader xssfReader = new XSSFReader(opcPkg);
            StylesTable styles = xssfReader.getStylesTable();

            // Content Handler
            ReadOnlySharedStringsTable ssTable = new ReadOnlySharedStringsTable(opcPkg);
            XSSFSheetContentsHandler<EB> sheetContentsHandler = new XSSFSheetContentsHandler<EB>(beanClz);
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

            return sheetContentsHandler.getRowObjects();
        } catch (Exception ex) {
            String errMsg =
                    String.format("Error reading sheet %d, ExcelBean %s : %s", sheetNo, beanClz, ex.getMessage());
            LOGGER.error(errMsg, ex);
            throw new ExcelReadException(errMsg, ex);
        }

    }



}
