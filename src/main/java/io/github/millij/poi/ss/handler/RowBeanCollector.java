package io.github.millij.poi.ss.handler;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Objects;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;


/**
 * 
 */
public final class RowBeanCollector<T> implements RowListener<T> {

    private static final Logger LOGGER = LoggerFactory.getLogger(RowBeanCollector.class);


    // Properties

    private final List<T> beans;


    // Constructors

    public RowBeanCollector() {
        super();

        // init
        this.beans = new ArrayList<>();
    }


    // Getters and Setters

    public List<T> getBeans() {
        return Collections.unmodifiableList(beans);
    }


    // RowListener
    // ------------------------------------------------------------------------

    @Override
    public void row(int rowNum, T rowObj) {
        if (Objects.isNull(rowObj)) {
            LOGGER.warn("NULL object returned for row : {}", rowNum);
            return;
        }

        beans.add(rowObj);
    }


}
