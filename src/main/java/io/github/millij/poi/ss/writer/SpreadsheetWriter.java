package io.github.millij.poi.ss.writer;

import java.io.File;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.logging.log4j.util.Strings;


/**
 * Representation of a Spreadsheet Writer. To write data to files, create a new Instance of {@link SpreadsheetWriter},
 * add data to sheets, and finally write the workbook to file.
 * 
 * @since 3.0
 */
public interface SpreadsheetWriter {

    //
    // Add Sheet :: Beans

    /**
     * Add a new sheet and add the beans a rows of data. To write the data to a file, call {@link #write(String)} method
     * once the data addition is completed.
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
     * Add a new sheet and add the beans a rows of data. To write the data to a file, call {@link #write(String)} method
     * once the data addition is completed.
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
     * Add a new sheet and add the beans a rows of data. To write the data to a file, call {@link #write(String)} method
     * once the data addition is completed.
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
     * Add a new sheet and add the beans a rows of data. To write the data to a file, call {@link #write(String)} method
     * once the data addition is completed.
     * 
     * @param beanClz The Class type to serialize the rows data
     * @param beans List of Data beans of the parameterized type
     * @param sheetName Name of the Sheet. (set it to <code>null</code> for default name)
     * @param headers a {@link List} of Header names to write in the file. <code>null</code> or empty list will default
     *        to all writable properties.
     */
    <T> void addSheet(Class<T> beanClz, List<T> beans, String sheetName, List<String> headers);


    //
    // Add Sheet :: Map<String, Object>

    /**
     * Add a new sheet and add the rows of data defined by a {@link Map}. The <code>null</code> entry in the rows data
     * list will be skipped and no rows will be added to the sheet.
     * 
     * @param rowsData List of Rows defined by a {@link Map}s where the map key is header and value is the row value
     * @param headers a {@link List} of Header names to write in the file
     */
    default void addSheet(List<Map<String, Object>> rowsData, List<String> headers) {
        this.addSheet(rowsData, null, headers);
    }

    /**
     * Add a new sheet and add the rows of data defined by a {@link Map}. The <code>null</code> entry in the rows data
     * list will be skipped and no rows will be added to the sheet.
     * 
     * @param rowsData List of Rows defined by a {@link Map}s where the map key is header and value is the row value
     * @param sheetName Name of the Sheet. (set it to <code>null</code> for default name)
     * @param headers a {@link List} of Header names to write in the file
     */
    void addSheet(List<Map<String, Object>> rowsData, String sheetName, List<String> headers);


    //
    // Write

    /**
     * Writes the current Spreadsheet workbook to a file in the specified path.
     * 
     * @param filepath output filepath (including filename).
     * 
     * @throws IOException if the filepath is not writable.
     */
    default void write(final String filepath) throws IOException {
        // Sanity checks
        if (Strings.isBlank(filepath)) {
            throw new IllegalArgumentException("#write :: Input File Path is BLANK");
        }

        this.write(new File(filepath));
    }

    /**
     * Writes the current Spreadsheet workbook to a file
     * 
     * @param fiel output file
     * 
     * @throws IOException if the file is not writable.
     */
    void write(File file) throws IOException;


}
