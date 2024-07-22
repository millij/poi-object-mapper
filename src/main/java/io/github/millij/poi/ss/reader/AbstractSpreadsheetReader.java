package io.github.millij.poi.ss.reader;

import java.io.InputStream;
import java.util.List;
import java.util.Map;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import io.github.millij.poi.SpreadsheetReadException;
import io.github.millij.poi.ss.handler.RowBeanCollector;


/**
 * A abstract implementation of {@link SpreadsheetReader}.
 */
abstract class AbstractSpreadsheetReader implements SpreadsheetReader {

    private static final Logger LOGGER = LoggerFactory.getLogger(AbstractSpreadsheetReader.class);


    protected final int headerRowIdx;
    protected final int lastRowIdx;


    // Constructor
    // ------------------------------------------------------------------------

    AbstractSpreadsheetReader(final int headerRowIdx, final int lastRowIdx) {
        super();

        // init
        this.headerRowIdx = headerRowIdx;
        this.lastRowIdx = lastRowIdx;

        final String insName = this.getClass().getSimpleName();
        LOGGER.debug("Successfully instantiated {} : header #{}, lastRow #{}", insName, headerRowIdx, lastRowIdx);
    }


    // Abstract Methods
    // ------------------------------------------------------------------------


    // Methods
    // ------------------------------------------------------------------------

    @Override
    public <T> List<T> read(final Class<T> beanClz, final InputStream is) throws SpreadsheetReadException {
        // Row Collector
        final RowBeanCollector<T> beanCollector = new RowBeanCollector<>();

        // Read with callback to fill list
        this.read(beanClz, is, beanCollector);

        // Result
        final List<T> beans = beanCollector.getBeans();
        return beans;
    }

    @Override
    public <T> List<T> read(final Class<T> beanClz, final InputStream is, final int sheetNo)
            throws SpreadsheetReadException {
        // Row Collector
        final RowBeanCollector<T> beanCollector = new RowBeanCollector<>();

        // Read with callback to fill list
        this.read(beanClz, is, sheetNo, beanCollector);

        // Result
        final List<T> beans = beanCollector.getBeans();
        return beans;
    }


    //
    // Read to Map

    @Override
    public List<Map<String, Object>> read(final InputStream is) throws SpreadsheetReadException {
        // Row Collector
        final RowBeanCollector<Map<String, Object>> beanCollector = new RowBeanCollector<>();

        // Read with callback to fill list
        this.read(is, beanCollector);

        // Result
        final List<Map<String, Object>> beans = beanCollector.getBeans();
        return beans;
    }

    @Override
    public List<Map<String, Object>> read(final InputStream is, final int sheetNo) throws SpreadsheetReadException {
        // Row Collector
        final RowBeanCollector<Map<String, Object>> beanCollector = new RowBeanCollector<>();

        // Read with callback to fill list
        this.read(is, sheetNo, beanCollector);

        // Result
        final List<Map<String, Object>> beans = beanCollector.getBeans();
        return beans;
    }


}
