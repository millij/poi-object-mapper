package io.github.millij.poi.ss.writer;

import java.io.IOException;
import java.util.List;
import java.util.Map;


/**
 * Representation of a Spreadsheet Writer. To write data to files, create a new Instance of {@link SpreadsheetWriter},
 * add data to sheets, and finally write the workbook to file.
 * 
 * @since 3.0
 */
public interface SpreadsheetWriter {

    /**
     * This method will attempt to add a new sheet and add the beans a rows of data. To write the data to a file, call
     * {@link #write(String)} method once the data addition is completed.
     * 
     * <br/>
     * <b>Sheet Name</b> : default sheet name will be used
     * 
     * <br/>
     * <b>Headers</b> : all possible properties as defined by the Type will be used as headers
     * 
     * @param beanClz The Class type to serialize the rows data
     * @param beans List of Data beans of the parameterized type
     */
    default <T> void addSheet(Class<T> beanClz, List<T> beans) {
        this.addSheet(beanClz, beans, null, null);
    }

    /**
     * This method will attempt to add a new sheet and add the beans a rows of data. To write the data to a file, call
     * {@link #write(String)} method once the data addition is completed.
     * 
     * <br/>
     * <b>Headers</b> : all possible properties as defined by the Type will be used as headers
     * 
     * @param beanClz The Class type to serialize the rows data
     * @param beans List of Data beans of the parameterized type
     * @param sheetName Name of the Sheet. (set it to <code>null</code> for default name)
     */
    default <T> void addSheet(Class<T> beanClz, List<T> beans, String sheetName) {
        this.addSheet(beanClz, beans, sheetName, null);
    }

    /**
     * This method will attempt to add a new sheet and add the beans a rows of data. To write the data to a file, call
     * {@link #write(String)} method once the data addition is completed.
     * 
     * <br/>
     * <b>Sheet Name</b> : default sheet name will be used
     * 
     * @param beanClz The Class type to serialize the rows data
     * @param beans List of Data beans of the parameterized type
     * @param headers a {@link List} of Header names to write in the file. <code>null</code> or empty list will default
     *        to all writable properties.
     */
    default <T> void addSheet(Class<T> beanClz, List<T> beans, List<String> headers) {
        this.addSheet(beanClz, beans, null, headers);
    }

    /**
     * This method will attempt to add a new sheet and add the beans a rows of data. To write the data to a file, call
     * {@link #write(String)} method once the data addition is completed.
     * 
     * @param beanClz The Class type to serialize the rows data
     * @param beans List of Data beans of the parameterized type
     * @param sheetName Name of the Sheet. (set it to <code>null</code> for default name)
     * @param headers a {@link List} of Header names to write in the file. <code>null</code> or empty list will default
     *        to all writable properties.
     */
    <T> void addSheet(Class<T> beanClz, List<T> beans, String sheetName, List<String> headers);


    /**
     * This method will attempt to add a new sheet and add the rows of data from the rows data. The
     * <code>null</code> entry in the rows data list will be skipped and no rows will be added to the
     * sheet.
     * 
     * @param sheetName Name of the Sheet. (set it to <code>null</code> for default name)
     * @param headers a {@link List} of Header names to write in the file. <code>null</code> or empty
     *        list will default to all writable properties.
     * @param rowsData List of Map, representing the sheet data. Each Map represents the row data where
     *        the map elements stores a value for a header as key.
     */
    void addSheet(String sheetName, List<String> headers, List<Map<String, String>> rowsData);


    /**
     * Writes the current Spreadsheet workbook to a file in the specified path.
     * 
     * @param filepath output filepath (including filename).
     * 
     * @throws IOException if the filepath is not writable.
     */
    void write(String filepath) throws IOException;

}
