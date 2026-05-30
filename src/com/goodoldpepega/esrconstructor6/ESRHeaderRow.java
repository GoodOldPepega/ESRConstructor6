package com.goodoldpepega.esrconstructor6;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Anton Nikitin
 * 27.02.2023
 */


public class ESRHeaderRow implements ESRRow {
    private List<ESRCell> cellsList = new ArrayList<ESRCell>();
    int maxSize = 256;

    /**
     * Get header cells of header row
     * @return
     */
    public List<ESRCell> getCells() throws ESRException {
        if (cellsList.size() == maxSize) {
            throw new ESRException(ESRHeaderCell.class.getName(),
                    "EXCEL does not allows add more than " + maxSize + " columns");
        }
        return cellsList;
    }

    /**
     * Add header cell to header row.
     * Notice that EXCEL allows only 16384 columns in one sheet.
     * @param cell Header cell
     * @throws ESRException
     */
    public void addCell(ESRCell cell) throws ESRException {
        if (!(cell instanceof ESRHeaderCell)) {
            throw new ESRException(ESRHeaderCell.class.getName(),
                    "You can only add ESRHeaderCell to ESRHeaderRow.");
        }
        cellsList.add(cell);
    }

    public void setCellsList(List<ESRCell> cellsList) {
        this.cellsList = cellsList;
    }
}
