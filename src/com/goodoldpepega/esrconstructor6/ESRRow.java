package com.goodoldpepega.esrconstructor6;

import java.util.List;

/**
 * Report row interface
 */
public interface ESRRow {

    /**
     * Get cells of row
     * @return
     */
    List<ESRCell> getCells() throws ESRException;

    /**
     * Add cell to row
     * @param cell Cell
     * @throws ESRException
     */
    void addCell(ESRCell cell) throws ESRException;
}
