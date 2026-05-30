package com.goodoldpepega.esrconstructor6;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.ArrayList;
import java.util.List;

/**
 * Block of title
 * Created by Anton Nikitin
 * 04.03.2023
 */


public class ESRTitleBlock {
    private List<String> listOfTitleRows = new ArrayList<String>();
    private CellStyle cellStyle;

    /**
     * Get rows of title
     * @return
     */
    public List<String> getTitleRows() {
        return listOfTitleRows;
    }

    /**
     * Add row to title
     * @param titleRow
     */
    public void addTitleRow(String titleRow) {
        listOfTitleRows.add(titleRow);
    }

    /**
     * Get style of title cell
     * @return
     */
    public CellStyle getStyle() {
        return cellStyle;
    }

    /**
     * Set style of title cell
     * @param cellStyle
     */
    public void setStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }
}
