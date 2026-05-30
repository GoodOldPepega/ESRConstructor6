package com.goodoldpepega.esrconstructor6;

/**
 * Created by Anton Nikitin
 * 30.05.2026
 */


public class ESRTableRowS extends ESRTableRow {

    public ESRTableRowS(ESRTableRow esrTableRow) throws ESRException {
        maxSize = 16384;
        this.setCellsList(esrTableRow.getCells());
    }
}
