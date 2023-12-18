package io.github.millij.poi.ss.reader;

import java.io.InputStream;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

import io.github.millij.poi.SpreadsheetReadException;
import io.github.millij.poi.ss.handler.RowContentsHandler;
import io.github.millij.poi.ss.handler.RowListener;
import io.github.millij.poi.util.Beans;


/**
 * Reader implementation of {@link Workbook} for an OOXML .xlsx file. This implementation is suitable for low memory SAX
 * parsing or similar.
 * 
 * @see XlsReader
 */
public class XlsxReader extends AbstractSpreadsheetReader {

    private static final Logger LOGGER = LoggerFactory.getLogger(XlsxReader.class);


    // Constructor

    public XlsxReader() {
        this(0);
    }

    public XlsxReader(final int headerRowIdx) {
        this(headerRowIdx, Integer.MAX_VALUE);
    }

    public XlsxReader(final int headerRowIdx, final int lastRowIdx) {
        super(headerRowIdx, lastRowIdx);
    }


    // SpreadsheetReader Impl
    // ------------------------------------------------------------------------

    @Override
    public <T> void read(final Class<T> beanClz, final InputStream is, final RowListener<T> listener)
            throws SpreadsheetReadException {
        // Sanity checks
        if (!Beans.isInstantiableType(beanClz)) {
            throw new IllegalArgumentException("XlsxReader :: Invalid bean type passed !");
        }

        try (final OPCPackage opcPkg = OPCPackage.open(is)) {
            // XSSF Reader
            final XSSFReader xssfReader = new XSSFReader(opcPkg);

            // Content Handler
            final StylesTable styles = xssfReader.getStylesTable();
            final ReadOnlySharedStringsTable ssTable = new ReadOnlySharedStringsTable(opcPkg);
            final SheetContentsHandler sheetHandler =
                    new RowContentsHandler<T>(beanClz, listener, headerRowIdx, lastRowIdx);

            final ContentHandler handler = new XSSFSheetXMLHandler(styles, ssTable, sheetHandler, true);

            // XML Reader
            final XMLReader xmlParser = XMLHelper.newXMLReader();
            xmlParser.setContentHandler(handler);

            // Iterate over sheets
            for (SheetIterator worksheets = (SheetIterator) xssfReader.getSheetsData(); worksheets.hasNext();) {
                final InputStream sheetInpStream = worksheets.next();
                final String sheetName = worksheets.getSheetName();
                LOGGER.debug("Reading XLSX Sheet :: ", sheetName);

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
    public <T> void read(final Class<T> beanClz, final InputStream is, final int sheetNo, RowListener<T> listener)
            throws SpreadsheetReadException {
        // Sanity checks
        if (!Beans.isInstantiableType(beanClz)) {
            throw new IllegalArgumentException("XlsxReader :: Invalid bean type passed !");
        }

        try (final OPCPackage opcPkg = OPCPackage.open(is)) {
            // XSSF Reader
            final XSSFReader xssfReader = new XSSFReader(opcPkg);

            // Content Handler
            final StylesTable styles = xssfReader.getStylesTable();
            final ReadOnlySharedStringsTable ssTable = new ReadOnlySharedStringsTable(opcPkg);
            final SheetContentsHandler sheetHandler =
                    new RowContentsHandler<T>(beanClz, listener, headerRowIdx, lastRowIdx);

            final ContentHandler handler = new XSSFSheetXMLHandler(styles, ssTable, sheetHandler, true);

            // XML Reader
            final XMLReader xmlParser = XMLHelper.newXMLReader();
            xmlParser.setContentHandler(handler);

            // Iterate over sheets
            final SheetIterator worksheets = (SheetIterator) xssfReader.getSheetsData();
            for (int i = 1; worksheets.hasNext(); i++) {
                // Get Sheet
                final InputStream sheetInpStream = worksheets.next();
                final String sheetName = worksheets.getSheetName();

                // Check for Sheet No.
                if (i != sheetNo) {
                    continue;
                }

                LOGGER.debug("Reading XLSX Sheet :: #{} - {}", i, sheetName);

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
