package com.goodoldpepega.esrconstructor6;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

/**
 * Created by Anton Nikitin
 * 01.03.2023
 */


public class ESRStyleHandler_old {

    private Workbook workbook;
    CellStyle headerCellStyle;
    CellStyle titleCellStyle;
    CellStyle defaultCellStyle;
    CellStyle greenCellStyle;
    CellStyle yellowCellStyle;
    CellStyle redCellStyle;
    CellStyle strippedCellStyle1;
    CellStyle strippedCellStyle2;

    public ESRStyleHandler_old(Workbook workbook) {
        this.workbook = workbook;
        initCellStyles();

    }

    private void initCellStyles() {
        Color customGreenColor;
        Color customYellowColor;
        Color customRedColor;
        Color customStripped1Color;
        Color customStripped2Color;

        headerCellStyle = workbook.createCellStyle();
        titleCellStyle = workbook.createCellStyle();
        defaultCellStyle = workbook.createCellStyle();
        greenCellStyle = workbook.createCellStyle();
        yellowCellStyle = workbook.createCellStyle();
        redCellStyle = workbook.createCellStyle();
        strippedCellStyle1 = workbook.createCellStyle();
        strippedCellStyle2 = workbook.createCellStyle();

        if (workbook instanceof HSSFWorkbook) {
            HSSFPalette palette = ((HSSFWorkbook) workbook).getCustomPalette();
            palette.setColorAtIndex(HSSFColor.INDIGO.index, (byte)90,(byte)200,(byte)15);
            palette.setColorAtIndex(HSSFColor.LAVENDER.index, (byte)250,(byte)230,(byte)40);
            palette.setColorAtIndex(HSSFColor.LIME.index, (byte)250,(byte)30,(byte)5);
            palette.setColorAtIndex(HSSFColor.VIOLET.index, (byte)255,(byte)250,(byte)230);
            palette.setColorAtIndex(HSSFColor.OLIVE_GREEN.index, (byte)255,(byte)235,(byte)210);
            customGreenColor = palette.getColor(HSSFColor.INDIGO.index);
            customYellowColor = palette.getColor(HSSFColor.LAVENDER.index);
            customRedColor = palette.getColor(HSSFColor.LIME.index);
            customStripped1Color = palette.getColor(HSSFColor.VIOLET.index);
            customStripped2Color = palette.getColor(HSSFColor.OLIVE_GREEN.index);
        }
        else {
            XSSFColor xssfColor = new XSSFColor(new byte[]{(byte) 90, (byte) 200, (byte) 15}, new DefaultIndexedColorMap());
            customGreenColor = xssfColor;
            xssfColor = new XSSFColor(new byte[]{(byte) 250, (byte) 230, (byte) 40}, new DefaultIndexedColorMap());
            customYellowColor = xssfColor;
            xssfColor = new XSSFColor(new byte[]{(byte) 250, (byte) 30, (byte) 5}, new DefaultIndexedColorMap());
            customRedColor = xssfColor;
            xssfColor = new XSSFColor(new byte[]{(byte) 255, (byte) 255, (byte) 230}, new DefaultIndexedColorMap());
            customStripped1Color = xssfColor;
            xssfColor = new XSSFColor(new byte[]{(byte) 255, (byte) 235, (byte) 210}, new DefaultIndexedColorMap());
            customStripped2Color = xssfColor;
        }
        DataFormat dataFormat = workbook.createDataFormat();

        CellStyle[] defaultStylesArray = new CellStyle[]{headerCellStyle, titleCellStyle, defaultCellStyle,
                greenCellStyle, yellowCellStyle, redCellStyle, strippedCellStyle1,
                strippedCellStyle2};

        for (CellStyle cellStyle : defaultStylesArray) {
            cellStyle.setWrapText(true);
            cellStyle.setBorderBottom(BorderStyle.MEDIUM);
            cellStyle.setBorderTop(BorderStyle.MEDIUM);
            cellStyle.setBorderLeft(BorderStyle.MEDIUM);
            cellStyle.setBorderRight(BorderStyle.MEDIUM);

        }

        Font hdrFont = workbook.createFont();
        hdrFont.setBold(true);
        headerCellStyle.setFont(hdrFont);
        headerCellStyle.setAlignment(HorizontalAlignment.CENTER);
        headerCellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);


        titleCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCellStyle.setFont(hdrFont);


        defaultCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);


        greenCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        greenCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);


        yellowCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        yellowCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        redCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        redCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);


        strippedCellStyle1.setVerticalAlignment(VerticalAlignment.CENTER);
        strippedCellStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        strippedCellStyle2.setVerticalAlignment(VerticalAlignment.CENTER);
        strippedCellStyle2.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        if (workbook instanceof HSSFWorkbook) {
            greenCellStyle.setFillForegroundColor(((HSSFColor)customGreenColor).getIndex());
            yellowCellStyle.setFillForegroundColor(((HSSFColor)customYellowColor).getIndex());
            redCellStyle.setFillForegroundColor(((HSSFColor)customRedColor).getIndex());
            strippedCellStyle1.setFillForegroundColor(((HSSFColor)customStripped1Color).getIndex());
            strippedCellStyle2.setFillForegroundColor(((HSSFColor)customStripped2Color).getIndex());
        }
        else {
            ((XSSFCellStyle)greenCellStyle).setFillForegroundColor(((XSSFColor)customGreenColor));
            ((XSSFCellStyle)yellowCellStyle).setFillForegroundColor(((XSSFColor)customYellowColor));
            ((XSSFCellStyle)redCellStyle).setFillForegroundColor(((XSSFColor)customRedColor));
            ((XSSFCellStyle)strippedCellStyle1).setFillForegroundColor(((XSSFColor)customStripped1Color));
            ((XSSFCellStyle)strippedCellStyle2).setFillForegroundColor(((XSSFColor)customStripped2Color));
        }
    }
}
