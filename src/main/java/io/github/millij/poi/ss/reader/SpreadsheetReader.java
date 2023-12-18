package io.github.millij.poi.ss.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import io.github.millij.poi.SpreadsheetReadException;
import io.github.millij.poi.ss.handler.RowListener;


/**
 * A simple representation of a Spreadsheet Reader. Any reader implementation (HSSF or XSSF) is expected to implement
 * and provide the below APIs.
 */
public interface SpreadsheetReader {


    //
    // Read with Custom RowListener
    // ------------------------------------------------------------------------

    /**
     * Reads the spreadsheet file to beans of the given type. This method will attempt to read all the available sheets
     * of the file and creates the objects of the passed type.
     * 
     * <p>
     * The {@link RowListener} implementation callback gets triggered after reading each Row. Best Suited for reading
     * Large files in restricted memory environments.
     * </p>
     * 
     * @param <T> The Parameterized bean Class.
     * @param beanClz The Class type to deserialize the rows data
     * @param is {@link InputStream} of the spreadsheet file
     * @param listener Custom {@link RowListener} implementation for row data callbacks.
     * 
     * @throws SpreadsheetReadException when the file data is not readable or row data to bean mapping failed.
     */
    <T> void read(Class<T> beanClz, InputStream is, RowListener<T> listener) throws SpreadsheetReadException;

    /**
     * Reads the spreadsheet file to beans of the given type. This method will attempt to read all the available sheets
     * of the file and creates the objects of the passed type.
     * 
     * <p>
     * The {@link RowListener} implementation callback gets triggered after reading each Row. Best Suited for reading
     * Large files in restricted memory environments.
     * </p>
     * 
     * @param <T> The Parameterized bean Class.
     * @param beanClz The Class type to deserialize the rows data
     * @param file {@link File} object of the spreadsheet file
     * @param listener Custom {@link RowListener} implementation for row data callbacks.
     * 
     * @throws SpreadsheetReadException when the file data is not readable or row data to bean mapping failed.
     */
    default <T> void read(Class<T> beanClz, File file, RowListener<T> listener) throws SpreadsheetReadException {
        // Closeable
        try (final InputStream fis = new FileInputStream(file)) {
            this.read(beanClz, fis, listener);
        } catch (IOException ex) {
            String errMsg = String.format("Failed to read file as Stream : %s", ex.getMessage());
            throw new SpreadsheetReadException(errMsg, ex);
        }
    }


    /**
     * Reads the spreadsheet file to beans of the given type. Note that only the requested sheet (sheet numbers are
     * indexed from 0) will be read.
     * 
     * <p>
     * The {@link RowListener} implementation callback gets triggered after reading each Row. Best Suited for reading
     * Large files in restricted memory environments.
     * </p>
     * 
     * @param <T> The Parameterized bean Class.
     * @param beanClz The Class type to deserialize the rows data
     * @param is {@link InputStream} of the spreadsheet file
     * @param sheetNo index of the Sheet to be read (index starts from 1)
     * @param listener Custom {@link RowListener} implementation for row data callbacks.
     * 
     * @throws SpreadsheetReadException when the file data is not readable or row data to bean mapping failed.
     */
    <T> void read(Class<T> beanClz, InputStream is, int sheetNo, RowListener<T> listener)
            throws SpreadsheetReadException;

    /**
     * Reads the spreadsheet file to beans of the given type. Note that only the requested sheet (sheet numbers are
     * indexed from 0) will be read.
     * 
     * <p>
     * The {@link RowListener} implementation callback gets triggered after reading each Row. Best Suited for reading
     * Large files in restricted memory environments.
     * </p>
     * 
     * @param <T> The Parameterized bean Class.
     * @param beanClz The Class type to deserialize the rows data
     * @param file {@link File} object of the spreadsheet file
     * @param sheetNo index of the Sheet to be read (index starts from 1)
     * @param listener Custom {@link RowListener} implementation for row data callbacks.
     * 
     * @throws SpreadsheetReadException when the file data is not readable or row data to bean mapping failed.
     */
    default <T> void read(Class<T> beanClz, File file, int sheetNo, RowListener<T> listener)
            throws SpreadsheetReadException {
        // Closeable
        try (final InputStream fis = new FileInputStream(file)) {
            this.read(beanClz, fis, sheetNo, listener);
        } catch (IOException ex) {
            String errMsg = String.format("Failed to read file as Stream : %s", ex.getMessage());
            throw new SpreadsheetReadException(errMsg, ex);
        }
    }


    //
    // Read with default RowListener
    // ------------------------------------------------------------------------

    /**
     * Reads the spreadsheet file to beans of the given type. This method will attempt to read all the available sheets
     * of the file and creates the objects of the passed type.
     * 
     * @param <T> The Parameterized bean Class.
     * @param beanClz The Class type to deserialize the rows data
     * @param is {@link InputStream} of the spreadsheet file
     * 
     * @return a {@link List} of objects of the parameterized type
     * 
     * @throws SpreadsheetReadException when the file data is not readable or row data to bean mapping failed.
     */
    <T> List<T> read(Class<T> beanClz, InputStream is) throws SpreadsheetReadException;

    /**
     * Reads the spreadsheet file to beans of the given type. This method will attempt to read all the available sheets
     * of the file and creates the objects of the passed type.
     * 
     * @param <T> The Parameterized bean Class.
     * @param beanClz The Class type to deserialize the rows data
     * @param file {@link File} object of the spreadsheet file
     * 
     * @return a {@link List} of objects of the parameterized type
     * 
     * @throws SpreadsheetReadException when the file data is not readable or row data to bean mapping failed.
     */
    default <T> List<T> read(Class<T> beanClz, File file) throws SpreadsheetReadException {
        // Closeable
        try (final InputStream fis = new FileInputStream(file)) {
            return this.read(beanClz, fis);
        } catch (IOException ex) {
            String errMsg = String.format("Failed to read file as Stream : %s", ex.getMessage());
            throw new SpreadsheetReadException(errMsg, ex);
        }
    }


    /**
     * Reads the spreadsheet file to beans of the given type. Note that only the requested sheet (sheet numbers are
     * indexed from 0) will be read.
     * 
     * @param <T> The Parameterized bean Class.
     * @param beanClz The Class type to deserialize the rows data
     * @param is {@link InputStream} of the spreadsheet file
     * @param sheetNo index of the Sheet to be read (index starts from 1)
     * 
     * @return a {@link List} of objects of the parameterized type
     * 
     * @throws SpreadsheetReadException when the file data is not readable or row data to bean mapping failed.
     */
    <T> List<T> read(Class<T> beanClz, InputStream is, int sheetNo) throws SpreadsheetReadException;

    /**
     * Reads the spreadsheet file to beans of the given type. Note that only the requested sheet (sheet numbers are
     * indexed from 0) will be read.
     * 
     * @param <T> The Parameterized bean Class.
     * @param beanClz The Class type to deserialize the rows data
     * @param file file {@link File} object of the spreadsheet file
     * @param sheetNo index of the Sheet to be read (index starts from 1)
     * 
     * @return a {@link List} of objects of the parameterized type
     * 
     * @throws SpreadsheetReadException when the file data is not readable or row data to bean mapping failed.
     */
    default <T> List<T> read(Class<T> beanClz, File file, int sheetNo) throws SpreadsheetReadException {
        // Closeable
        try (final InputStream fis = new FileInputStream(file)) {
            return this.read(beanClz, fis, sheetNo);
        } catch (IOException ex) {
            String errMsg = String.format("Failed to read file as Stream : %s", ex.getMessage());
            throw new SpreadsheetReadException(errMsg, ex);
        }
    }


}

