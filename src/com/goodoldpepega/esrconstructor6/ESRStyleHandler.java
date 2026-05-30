package com.goodoldpepega.esrconstructor6;

import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

/**
 * Handler of cells styles
 * Created by Anton Nikitin
 * 28.05.2026
 */


public class ESRStyleHandler {

    private Workbook workbook;
    private CellStyle headerCellStyle;
    private CellStyle titleCellStyle;
    private CellStyle defaultCellStyle;
    private CellStyle greenCellStyle;
    private CellStyle yellowCellStyle;
    private CellStyle redCellStyle;
    private CellStyle strippedCellStyle1;
    private CellStyle strippedCellStyle2;

    /**
     * Style of workbook
     * @param workbook Workbook
     */
    public ESRStyleHandler(Workbook workbook) {
        this.workbook = workbook;
        initCellStyles();

    }

    /**
     * Initialize all standard cell styles
     */
    private void initCellStyles() {

        Color customGreenColor;
        Color customYellowColor;
        Color customRedColor;
        Color customStripped1Color;
        Color customStripped2Color;

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

        headerCellStyle = workbook.createCellStyle();
        titleCellStyle = workbook.createCellStyle();
        defaultCellStyle = workbook.createCellStyle();
        greenCellStyle = workbook.createCellStyle();
        yellowCellStyle = workbook.createCellStyle();
        redCellStyle = workbook.createCellStyle();
        strippedCellStyle1 = workbook.createCellStyle();
        strippedCellStyle2 = workbook.createCellStyle();

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


        titleCellStyle.setDataFormat(dataFormat.getFormat("@"));
        titleCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCellStyle.setFont(hdrFont);


        defaultCellStyle.setDataFormat(dataFormat.getFormat("@"));
        titleCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);


        greenCellStyle.setDataFormat(dataFormat.getFormat("@"));
        greenCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        greenCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);


        yellowCellStyle.setDataFormat(dataFormat.getFormat("@"));
        yellowCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        yellowCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        redCellStyle.setDataFormat(dataFormat.getFormat("@"));
        redCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        redCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);


        strippedCellStyle1.setDataFormat(dataFormat.getFormat("@"));
        strippedCellStyle1.setVerticalAlignment(VerticalAlignment.CENTER);
        strippedCellStyle1.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        strippedCellStyle2.setDataFormat(dataFormat.getFormat("@"));
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

    CellStyle getHeaderCellStyle() {
        return headerCellStyle;
    }

    void setHeaderCellStyle(CellStyle headerCellStyle) {
        this.headerCellStyle = headerCellStyle;
    }

    CellStyle getTitleCellStyle() {
        return titleCellStyle;
    }

    void setTitleCellStyle(CellStyle titleCellStyle) {
        this.titleCellStyle = titleCellStyle;
    }

    CellStyle getDefaultCellStyle() {
        return defaultCellStyle;
    }

    void setDefaultCellStyle(CellStyle defaultCellStyle) {
        this.defaultCellStyle = defaultCellStyle;
    }

    CellStyle getGreenCellStyle() {
        return greenCellStyle;
    }

    void setGreenCellStyle(CellStyle greenCellStyle) {
        this.greenCellStyle = greenCellStyle;
    }

    CellStyle getYellowCellStyle() {
        return yellowCellStyle;
    }

    void setYellowCellStyle(CellStyle yellowCellStyle) {
        this.yellowCellStyle = yellowCellStyle;
    }

    CellStyle getRedCellStyle() {
        return redCellStyle;
    }

    void setRedCellStyle(CellStyle redCellStyle) {
        this.redCellStyle = redCellStyle;
    }

    CellStyle getStrippedCellStyle1() {
        return strippedCellStyle1;
    }

    void setStrippedCellStyle1(CellStyle strippedCellStyle1) {
        this.strippedCellStyle1 = strippedCellStyle1;
    }

    CellStyle getStrippedCellStyle2() {
        return strippedCellStyle2;
    }

    void setStrippedCellStyle2(CellStyle strippedCellStyle2) {
        this.strippedCellStyle2 = strippedCellStyle2;
    }
}
