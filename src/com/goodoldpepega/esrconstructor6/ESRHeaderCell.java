package com.goodoldpepega.esrconstructor6;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Created by Anton Nikitin
 * 27.02.2023
 */


public class ESRHeaderCell implements ESRCell {
    private Object value;
    private int width;
    private CellStyle cellStyle;

    /**
     * Create new header cell
     * @param value Text value of cell
     * @param width width of column in sheet
     */
    public ESRHeaderCell(Object value, int width) {
        this.value = value;
        this.width = width;
    }

    /**
     * Get value of cell
     * @return
     */
    public Object getValue() {
        return value;
    }

    /**
     * Set value of cell
     * @param object
     */
    public void setValue(Object object) {
        this.value = object;
    }

    /**
     * Get style of cell
     * @return
     */
    public CellStyle getStyle() {
        return cellStyle;
    }

    /**
     * Set style of cell
     * @param cellStyle
     */
    public void setStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    /**
     * Get width of column of cell
     * @return
     */
    public int getWidth() {
        return width;
    }

    /**
     * Set width of cell column
     * @param width
     */
    public void setWidth(int width) {
        this.width = width;
    }
}
