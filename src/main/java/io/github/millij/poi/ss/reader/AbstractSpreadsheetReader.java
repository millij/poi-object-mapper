package io.github.millij.poi.ss.reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import io.github.millij.poi.SpreadsheetReadException;
import io.github.millij.poi.ss.handler.RowListener;

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
	public <T> List<T> read(Class<T> beanClz, File file, String sheetName) throws SpreadsheetReadException {
		// Sanity Checks
		if (Objects.isNull(sheetName)) {
			String errMsg = "Failed to read the file with SheetName : NULL";
			throw new SpreadsheetReadException(errMsg);
		}
		final int sheetNo = this.getSheetNo(beanClz, file, sheetName);
		return this.read(beanClz, file, sheetNo);
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

}
