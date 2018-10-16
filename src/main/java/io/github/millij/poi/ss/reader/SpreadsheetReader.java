package io.github.millij.poi.ss.reader;

import io.github.millij.poi.SpreadsheetReadException;
import io.github.millij.poi.ss.handler.RowListener;

import java.io.File;
import java.io.InputStream;
import java.util.List;


/**
 * An Abstract representation of a Spreadsheet Reader. Any reader implementation (HSSF or XSSF) is
 * expected to implement and provide the below APIs.
 * 
 */
public interface SpreadsheetReader {


    // Read with row eventHandler

    <T> void read(Class<T> beanClz, File file, RowListener<T> eventHandler) throws SpreadsheetReadException;

    <T> void read(Class<T> beanClz, InputStream is, RowListener<T> eventHandler) throws SpreadsheetReadException;


    <T> void read(Class<T> beanClz, File file, int sheetNo, RowListener<T> eventHandler)
            throws SpreadsheetReadException;

    <T> void read(Class<T> beanClz, InputStream is, int sheetNo, RowListener<T> eventHandler)
            throws SpreadsheetReadException;



    // Read all

    <T> List<T> read(Class<T> beanClz, File file) throws SpreadsheetReadException;

    <T> List<T> read(Class<T> beanClz, InputStream is) throws SpreadsheetReadException;


    <T> List<T> read(Class<T> beanClz, File file, int sheetNo) throws SpreadsheetReadException;

    <T> List<T> read(Class<T> beanClz, InputStream is, int sheetNo) throws SpreadsheetReadException;



}
