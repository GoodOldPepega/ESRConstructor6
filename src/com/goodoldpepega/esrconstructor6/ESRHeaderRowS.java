package com.goodoldpepega.esrconstructor6;

/**
 * Created by Anton Nikitin
 * 30.05.2026
 */


public class ESRHeaderRowS extends ESRHeaderRow {

    public ESRHeaderRowS(ESRHeaderRow esrHeaderRow) throws ESRException {
        maxSize = 16384;
        this.setCellsList(esrHeaderRow.getCells());
    }
}
