package com.goodoldpepega.esrconstructor6;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Anton Nikitin
 * 27.02.2023
 */


public class ESRTableRow implements ESRRow {
    private List<ESRCell> cellsList = new ArrayList<ESRCell>();
    int maxSize = 256;

    /**
     * Get cells of row
     * @return
     */
    public List<ESRCell> getCells() throws ESRException {
        if (cellsList.size() == maxSize) {
            throw new ESRException(ESRTableRow.class.getName(),
                    "EXCEL does not allows add more than " + maxSize + " columns");
        }
        return cellsList;
    }

    /**
     * Add cell to row.
     * @param cell Cell
     * @throws ESRException
     */
    public void addCell(ESRCell cell) throws ESRException {
        if (!(cell instanceof ESRTableCell)) {
            throw new ESRException(ESRTableRow.class.getName(),
                    "You can add only add ESRTableCell to ESRTableRow.");
        }
        cellsList.add(cell);
    }

    public void setCellsList(List<ESRCell> cellsList) {
        this.cellsList = cellsList;
    }
}
