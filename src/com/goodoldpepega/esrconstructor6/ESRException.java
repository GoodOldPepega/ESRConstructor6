package com.goodoldpepega.esrconstructor6;

/**
 * Created by Anton Nikitin
 * 04.03.2023
 */


public class ESRException extends Exception {

    public ESRException (String classNameStr, String exceptionDetails) {
        super(classNameStr + " : " + exceptionDetails);
    }
}
