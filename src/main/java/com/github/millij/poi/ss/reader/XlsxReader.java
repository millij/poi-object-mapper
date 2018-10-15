package com.github.millij.poi.ss.reader;

import static com.github.millij.poi.util.Beans.isInstantiableType;

import java.io.InputStream;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import com.github.millij.poi.SpreadsheetReadException;
import com.github.millij.poi.ss.handler.RowContentsHandler;
import com.github.millij.poi.ss.handler.RowListener;


/**
 * Reader impletementation of {@link Workbook} for an OOXML .xlsx file. This implementation is
 * suitable for low memory sax parsing or similar.
 * 
 * @see XlsReader
 */
public class XlsxReader extends AbstractSpreadsheetReader {

    private static final Logger LOGGER = LoggerFactory.getLogger(XlsxReader.class);


    // Constructor

    public XlsxReader() {
        super();
    }


    // SpreadsheetReader Impl
    // ------------------------------------------------------------------------

    @Override
    public <T> void read(Class<T> beanClz, InputStream is, RowListener<T> listener)
            throws SpreadsheetReadException {
        // Sanity checks
        if (!isInstantiableType(beanClz)) {
            throw new IllegalArgumentException("XlsxReader :: Invalid bean type passed !");
        }

        try {
            final OPCPackage opcPkg = OPCPackage.open(is);

            // XSSF Reader
            XSSFReader xssfReader = new XSSFReader(opcPkg);
            StylesTable styles = xssfReader.getStylesTable();

            // Content Handler
            ReadOnlySharedStringsTable ssTable = new ReadOnlySharedStringsTable(opcPkg);
            SheetContentsHandler sheetContentsHandler = new RowContentsHandler<T>(beanClz, listener, 0);
            ContentHandler handler = new XSSFSheetXMLHandler(styles, ssTable, sheetContentsHandler, true);

            // XML Reader
            XMLReader xmlParser = SAXHelper.newXMLReader();
            xmlParser.setContentHandler(handler);

            // Iterate over sheets
            XSSFReader.SheetIterator worksheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            for (int i = 0; worksheets.hasNext(); i++) {
                InputStream sheetInpStream = worksheets.next();

                String sheetName = worksheets.getSheetName();

                // Parse sheet
                xmlParser.parse(new InputSource(sheetInpStream));
                sheetInpStream.close();
            }
        } catch (Exception ex) {
            String errMsg = String.format("Error reading sheet data, to Bean %s : %s", beanClz, ex.getMessage());
            LOGGER.error(errMsg, ex);
            throw new SpreadsheetReadException(errMsg, ex);
        }
    }


    @Override
    public <T> void read(Class<T> beanClz, InputStream is, int sheetNo, RowListener<T> listener)
            throws SpreadsheetReadException {
        // Sanity checks
        if (!isInstantiableType(beanClz)) {
            throw new IllegalArgumentException("XlsxReader :: Invalid bean type passed !");
        }

        try {
            final OPCPackage opcPkg = OPCPackage.open(is);

            // XSSF Reader
            XSSFReader xssfReader = new XSSFReader(opcPkg);
            StylesTable styles = xssfReader.getStylesTable();

            // Content Handler
            ReadOnlySharedStringsTable ssTable = new ReadOnlySharedStringsTable(opcPkg);
            SheetContentsHandler sheetContentsHandler = new RowContentsHandler<T>(beanClz, listener, 0);
            ContentHandler handler = new XSSFSheetXMLHandler(styles, ssTable, sheetContentsHandler, true);

            // XML Reader
            XMLReader xmlParser = SAXHelper.newXMLReader();
            xmlParser.setContentHandler(handler);

            // Iterate over sheets
            XSSFReader.SheetIterator worksheets = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            for (int i = 0; worksheets.hasNext(); i++) {
                InputStream sheetInpStream = worksheets.next();
                if (i != sheetNo) {
                    continue;
                }

                String sheetName = worksheets.getSheetName();

                // Parse Sheet
                xmlParser.parse(new InputSource(sheetInpStream));
                sheetInpStream.close();
            }

        } catch (Exception ex) {
            String errMsg = String.format("Error reading sheet %d, to Bean %s : %s", sheetNo, beanClz, ex.getMessage());
            LOGGER.error(errMsg, ex);
            throw new SpreadsheetReadException(errMsg, ex);
        }
    }



}
