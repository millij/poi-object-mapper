package io.github.millij.poi.ss.reader;

import io.github.millij.poi.SpreadsheetReadException;
import io.github.millij.poi.ss.handler.RowListener;
import io.github.millij.poi.ss.model.annotations.SheetColumn;
import io.github.millij.poi.util.Beans;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


/**
 * A abstract implementation of {@link SpreadsheetReader}.
 */
abstract class AbstractSpreadsheetReader implements SpreadsheetReader {

    private static final Logger LOGGER = LoggerFactory.getLogger(AbstractSpreadsheetReader.class);



    // Abstract Methods
    // ------------------------------------------------------------------------



    // Methods
    // ------------------------------------------------------------------------

    @Override
    public <T> void read(Class<T> beanClz, File file, RowListener<T> callback) throws SpreadsheetReadException {
        // Closeble
        try (InputStream fis = new FileInputStream(file)) {

            // chain
            this.read(beanClz, fis, callback);
        } catch (IOException ex) {
            String errMsg = String.format("Failed to read file as Stream : %s", ex.getMessage());
            throw new SpreadsheetReadException(errMsg, ex);
        }
    }


    @Override
    public <T> void read(Class<T> beanClz, File file, int sheetNo, RowListener<T> callback)
            throws SpreadsheetReadException {
        // Sanity checks
        try {
            InputStream fis = new FileInputStream(file);

            // chain
            this.read(beanClz, fis, sheetNo, callback);
        } catch (IOException ex) {
            String errMsg = String.format("Failed to read file as Stream : %s", ex.getMessage());
            throw new SpreadsheetReadException(errMsg, ex);
        }
    }



    @Override
    public <T> List<T> read(Class<T> beanClz, File file) throws SpreadsheetReadException {
        // Closeble
        try (InputStream fis = new FileInputStream(file)) {
            return this.read(beanClz, fis);
        } catch (IOException ex) {
            String errMsg = String.format("Failed to read file as Stream : %s", ex.getMessage());
            throw new SpreadsheetReadException(errMsg, ex);
        }
    }

    @Override
    public <T> List<T> read(Class<T> beanClz, InputStream is) throws SpreadsheetReadException {
        // Result
        final List<T> sheetBeans = new ArrayList<T>();

        // Read with callback to fill list
        this.read(beanClz, is, new RowListener<T>() {

            @Override
            public void row(int rowNum, T rowObj) {
                if (rowObj == null) {
                    LOGGER.error("Null object returned for row : {}", rowNum);
                    return;
                }

                sheetBeans.add(rowObj);
            }

        });

        return sheetBeans;
    }


    @Override
    public <T> List<T> read(Class<T> beanClz, File file, int sheetNo) throws SpreadsheetReadException {
        // Closeble
        try (InputStream fis = new FileInputStream(file)) {
            return this.read(beanClz, fis, sheetNo);
        } catch (IOException ex) {
            String errMsg = String.format("Failed to read file as Stream : %s", ex.getMessage());
            throw new SpreadsheetReadException(errMsg, ex);
        }
    }

    @Override
    public <T> List<T> read(Class<T> beanClz, InputStream is, int sheetNo) throws SpreadsheetReadException {
        // Result
        final List<T> sheetBeans = new ArrayList<T>();

        // Read with callback to fill list
        this.read(beanClz, is, sheetNo, new RowListener<T>() {

            @Override
            public void row(int rowNum, T rowObj) {
                if (rowObj == null) {
                    LOGGER.error("Null object returned for row : {}", rowNum);
                    return;
                }

                sheetBeans.add(rowObj);
            }

        });

        return sheetBeans;
    }
    
    public static String getReturnType(Class<?> beanClz, String header) {


        String headerType = null;
        // Fields
        Field[] fields = beanClz.getDeclaredFields();
        for (Field f : fields) {
            if (!f.isAnnotationPresent(SheetColumn.class)) {
                continue;
            }
            String fieldName = f.getName();
            SheetColumn ec = f.getDeclaredAnnotation(SheetColumn.class);
            if (header.equals(fieldName) || header.equals(ec.value())) {
                headerType = f.getType().getName();
            }
            continue;
        }

        // Methods
        Method[] methods = beanClz.getDeclaredMethods();
        for (Method m : methods) {
            if (!m.isAnnotationPresent(SheetColumn.class)) {
                continue;
            }
            String fieldName = Beans.getFieldName(m);
            SheetColumn ec = m.getDeclaredAnnotation(SheetColumn.class);
            if (header.equals(fieldName) || header.equals(ec.value())) {
                headerType = m.getReturnType().getName();
            }
            continue;
        }
        if (StringUtils.isBlank(headerType)) {
            LOGGER.info("Failed to get the return type of the given Header '{}'", header);
        }
        return headerType;

    }


}
