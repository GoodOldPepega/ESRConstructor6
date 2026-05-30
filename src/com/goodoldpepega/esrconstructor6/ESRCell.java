package com.goodoldpepega.esrconstructor6;

import org.apache.poi.ss.usermodel.CellStyle;

public interface ESRCell {

    int CELL_TYPE_NUMERIC = 0;
    int CELL_TYPE_STRING = 1;
    int CELL_TYPE_FORMULA = 2;
    int CELL_TYPE_BLANK = 3;
    int CELL_TYPE_BOOLEAN = 4;

    /**
     * Get cell value
     * @return
     */
    Object getValue();

    /**
     * Set cell value
     * @param object
     */
    void setValue(Object object);

    /**
     * Get cell style
     * @return
     */
    public CellStyle getStyle();

    /**
     * Set cell style
     * @param hssfCellStyle
     */
    public void setStyle(CellStyle hssfCellStyle);
}
