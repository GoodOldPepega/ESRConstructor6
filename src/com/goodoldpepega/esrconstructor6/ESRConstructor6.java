package com.goodoldpepega.esrconstructor6;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Anton Nikitin
 * 28.05.2026
 */


public class ESRConstructor6 {

    private Workbook workbook;
    private ESRStyleHandler esrStyleHandler;
    private CellStyle currentStripeColor;
    private int rowIndex;
    private int stripeColorNumber;
    private boolean hasHeader;
    private boolean widthComplete;
    private Map<Integer, Integer> mapOfWidth;
    private static int lastRowIndex;
    private boolean isNewFormat = true;

    public ESRConstructor6(Workbook workbook) {
        this.workbook = workbook;
        if (workbook instanceof HSSFWorkbook){
            isNewFormat = false;
            lastRowIndex = 65537;
        } else {
            lastRowIndex = 1048576;
        }
        esrStyleHandler = new ESRStyleHandler(workbook);
    }

    private void reset(){
        rowIndex = 0;
        stripeColorNumber = 0;
        hasHeader = true;
        widthComplete = false;
        resetStripeColor();
    }
    private void resetStripeColor(){
        currentStripeColor = esrStyleHandler.getStrippedCellStyle1();
    }

    public void writeToExcel(ESRReport report, String sheetName, boolean needStrippedColourScheme) throws ESRException {
        reset();
        ESRTitleBlock titleBlock = report.getEsrTitleBlock();
        ESRHeaderBlock headerBlock = report.getEsrHeaderBlock();
        List<ESRBlock> listOfValueBlocks = report.getEsrBlocksList();
        int sheetIndex = 1;
        Sheet sheet = createSheet(handleSheetName(sheetName, sheetIndex));
        int titleRowsMergeQuantity = 1;
        int headerRowIndex = 0;
        if (headerBlock != null){
            titleRowsMergeQuantity = headerBlock.getMaxHeaderRowLength();
            if (titleBlock != null){
                headerRowIndex = titleBlock.getTitleRows().size();
            }
            createHeader(headerBlock, sheet, headerRowIndex);
        } else {
            hasHeader = false;
            mapOfWidth = new HashMap<Integer, Integer>();
            if (listOfValueBlocks.size() > 0) {
                for (ESRBlock rowsBlock : listOfValueBlocks) {
                    List<ESRRow> rows = rowsBlock.getRows();
                    if (rows.size() > 0) {
                        ESRRow esrRow = rows.get(0);
                        titleRowsMergeQuantity = esrRow.getCells().size();
                    }
                }
            }
        }
        if (titleBlock != null){
            if (headerBlock == null){
                rowIndex = titleBlock.getTitleRows().size();
            }
            createTitleRows(titleBlock, sheet, titleRowsMergeQuantity);
        }
        for (int i = 0; i < listOfValueBlocks.size(); i++) {
            ESRBlock rowsBlock = listOfValueBlocks.get(i);
            if (rowIndex + rowsBlock.getRows().size() >= lastRowIndex) {
                if (sheetIndex == 1) {
                    workbook.setSheetName(workbook.getSheetIndex(sheet), sheetName + "_" + sheetIndex);
                }
                sheet = createSheet(handleSheetName(sheetName, ++sheetIndex));
                if (titleBlock != null){
                    if (headerBlock == null){
                        rowIndex = titleBlock.getTitleRows().size();
                    }
                    createTitleRows(titleBlock, sheet, titleRowsMergeQuantity);
                }
                if (headerBlock != null) {
                    createHeader(headerBlock, sheet, headerRowIndex);
                }
                if (titleBlock == null && headerBlock == null){
                    rowIndex = 0;
                }
                resetStripeColor();
            }
            writeRowsBlock(rowsBlock, sheet, needStrippedColourScheme);
            toggleStripeColors();
        }
        if (!hasHeader) {
            for (Integer columnIndex : mapOfWidth.keySet()) {
                sheet.setColumnWidth(columnIndex, mapOfWidth.get(columnIndex));
            }
        }
    }

    private String handleSheetName(String name, int sheetIndex){
        if (sheetIndex > 1){
            name += "_" + sheetIndex;
        }
        if (name.length() > 28){
            name = name.substring(0, 28);
        }
        return name;
    }

    private Sheet createSheet (String sheetName) {
        return workbook.createSheet(sheetName);
    }

    private void createHeader(ESRHeaderBlock headerBlock, Sheet sheet, int headerRowIndex) throws ESRException {
        this.rowIndex = headerRowIndex;
        int firstRowIndex = rowIndex;
        List<ESRRow> headerRows = headerBlock.getRows();
        for (int k = 0; k < headerRows.size(); k++) {
            Object esrHeaderRowObj = headerRows.get(k);
            ESRRow esrHeaderRow;
            if (esrHeaderRowObj instanceof ESRHeaderRow){
                if (isNewFormat){
                    esrHeaderRow = new ESRHeaderRowS((ESRHeaderRow) headerRows.get(k));
                } else {
                    esrHeaderRow = headerRows.get(k);
                }
            } else {
                throw new ESRException(ESRHeaderCell.class.getName(),
                        "You can only add ESRHeaderCell to ESRHeaderRow.");
            }

            Row row = sheet.createRow(rowIndex);
            List<ESRCell> listOfCells = esrHeaderRow.getCells();
            for (int i = 0; i < listOfCells.size(); i++) {
                ESRHeaderCell esrHeaderCell = (ESRHeaderCell) listOfCells.get(i);
                sheet.setColumnWidth(i, esrHeaderCell.getWidth());
                Cell cell = row.createCell(i);
                if (esrHeaderCell.getStyle() == null) {
                    esrHeaderCell.setStyle(esrStyleHandler.getHeaderCellStyle());
                }
                cell.setCellStyle(esrHeaderCell.getStyle());
                cell.setCellValue(String.valueOf(esrHeaderCell.getValue()));
            }
            rowIndex++;
        }
        if (headerBlock.getMerge() != null) {
            ESRMerge blockMerge = headerBlock.getMerge();
            List<int[]> listOfMerges = blockMerge.getListOfMerges();
            for (int[] mergeArray : listOfMerges) {
                try{
                    sheet.addMergedRegion(new CellRangeAddress(firstRowIndex + mergeArray[0],
                            firstRowIndex + mergeArray[1], mergeArray[2], mergeArray[3]));
                } catch (Exception e) {
                    System.out.println("ESRConstructor: Merge " + mergeArray[0] + mergeArray[1] + mergeArray[2] +
                            mergeArray[3] + " exceeds existing cells range in header block");
                }
            }
        }
    }

    private void createTitleRows(ESRTitleBlock titleBlock, Sheet sheet, int mergedColumnsNumber) {
        List<String> listOfTitleRows = titleBlock.getTitleRows();
        int titleRowIndex = 0;
        for (String titleString : listOfTitleRows) {
            Row row = sheet.createRow(titleRowIndex);
            for (int i = 0; i < mergedColumnsNumber; i++) {
                Cell cell = row.createCell(i);
                if (titleBlock.getStyle() == null) {
                    titleBlock.setStyle(esrStyleHandler.getTitleCellStyle());
                }
                cell.setCellStyle(titleBlock.getStyle());
                if (i == 0) {
                    cell.setCellValue(titleString);
                }
            }
            int newMergedColumnsNumber = mergedColumnsNumber;
            if (newMergedColumnsNumber > 0) {
                newMergedColumnsNumber--;
            }
            if (newMergedColumnsNumber != 0){
                sheet.addMergedRegion(new CellRangeAddress(titleRowIndex, titleRowIndex, 0,
                        newMergedColumnsNumber));
            }
            titleRowIndex++;
        }
    }

    private void writeRowsBlock(ESRBlock rowsBlock, Sheet sheet, boolean needStrippedColourScheme) throws ESRException {
        int firstRowIndex = rowIndex;
        boolean hasMerge = rowsBlock.getMerge() != null;
        List<ESRRow> listOfRows = rowsBlock.getRows();
        for (int i = 0; i < listOfRows.size(); i++) {
            ESRRow esrTableRow;
            if (isNewFormat){
                esrTableRow = new ESRTableRowS((ESRTableRow) listOfRows.get(i));
            } else {
                esrTableRow = listOfRows.get(i);
            }
            Row row = sheet.createRow(rowIndex++);
            writeRow(esrTableRow, row, needStrippedColourScheme, hasMerge);
        }
        if (rowsBlock.getMerge() != null) {
            ESRMerge esrMerge = rowsBlock.getMerge();
            for (int[] mergeArray : esrMerge.getListOfMerges()) {
                try{
                    sheet.addMergedRegionUnsafe(new CellRangeAddress(firstRowIndex + mergeArray[0],
                            firstRowIndex + mergeArray[1], mergeArray[2], mergeArray[3]));
                }
                catch (Exception e) {
                    System.out.println("ESRConstructor: Merge " + mergeArray[0] + mergeArray[1] + mergeArray[2] +
                            mergeArray[3] + " exceeds existing cells range in table block`");
                }
            }
        }
    }

    private void writeRow(ESRRow srRow, Row row, boolean needStrippedColourScheme, boolean hasMerge) throws ESRException {
        List<ESRCell> esrCells = srRow.getCells();
        int maxLinesQuantity = 0;
        int esrCellsSize = esrCells.size();
        for (int i = 0; i < esrCellsSize; i++) {
            ESRTableCell esrTableCell = (ESRTableCell) esrCells.get(i);
            if (!hasHeader && !widthComplete) {
                int width = esrTableCell.getWidth();
                if (width > 0) {
                    if (mapOfWidth.get(i) == null) {
                        mapOfWidth.put(i, width);
                        if (mapOfWidth.size() == esrCellsSize) {
                            widthComplete = true;
                        }
                    }
                }
            }
            Cell cell = row.createCell(i);
            CellStyle cellStyle;
            if (esrTableCell.getStyle() != null) {
                cellStyle = esrTableCell.getStyle();
            } else if (needStrippedColourScheme) {
                cellStyle = currentStripeColor;
            } else {
                cellStyle = esrStyleHandler.getDefaultCellStyle();
            }
            Object cellValueObject = esrTableCell.getValue();
            if (hasMerge) {
                int colWidth = row.getSheet().getColumnWidth(i);
                int valueLength = String.valueOf(cellValueObject).length();
                int newLineQuantity = String.valueOf(cellValueObject).split("\n").length;
                if (newLineQuantity > 0) {
                    newLineQuantity = newLineQuantity - 1;
                }
                int linesQuantity = (valueLength * 256 / colWidth) + newLineQuantity;
                if (maxLinesQuantity < linesQuantity) {
                    maxLinesQuantity = linesQuantity;
                }
            }
            setValueToCell(esrTableCell, cell);
            cell.setCellStyle(cellStyle);
            setHyperlink(esrTableCell, cell);
        }
        if (maxLinesQuantity == 0) {
            maxLinesQuantity = 1;
        }
        if (hasMerge) {
            row.setHeightInPoints(maxLinesQuantity * 10 * 1.7F);
        }
    }

    private void setValueToCell(ESRTableCell esrTableCell, Cell cell){
        Object cellValueObject = esrTableCell.getValue();
        String valueS = String.valueOf(cellValueObject);
        if (esrTableCell.getType() == Cell.CELL_TYPE_NUMERIC && ESRHelper.checkIfDouble(valueS)) {
            cell.setCellType(Cell.CELL_TYPE_NUMERIC);
            cell.setCellValue(Double.parseDouble(valueS));
        } else if (esrTableCell.getType() == Cell.CELL_TYPE_FORMULA) {
            cell.setCellType(Cell.CELL_TYPE_FORMULA);
            cell.setCellFormula(valueS);
        } else if (esrTableCell.getType() == Cell.CELL_TYPE_BLANK) {
            cell.setCellType(Cell.CELL_TYPE_BLANK);
        } else if (esrTableCell.getType() == Cell.CELL_TYPE_BOOLEAN) {
            if ("true".equals(valueS) || "false".equals(valueS)) {
                cell.setCellType(Cell.CELL_TYPE_BOOLEAN);
                cell.setCellValue(Boolean.parseBoolean(valueS));
            } else {
                cell.setCellValue(valueS);
            }
        } else if (esrTableCell.getType() == Cell.CELL_TYPE_ERROR) {
            cell.setCellType(Cell.CELL_TYPE_ERROR);
        } else {
            cell.setCellValue(valueS);
        }
    }

    private void setHyperlink(ESRTableCell esrTableCell, Cell cell){
        if (esrTableCell.getHyperlink() != null) {
            String url = esrTableCell.getHyperlink();
            Hyperlink link = (workbook).getCreationHelper().createHyperlink(HyperlinkType.URL);
            link.setAddress(url);
            cell.setHyperlink(link);
        }
    }

    private void toggleStripeColors(){
        stripeColorNumber++;
        if (stripeColorNumber % 2 == 0) {
            resetStripeColor();
        } else {
            currentStripeColor = esrStyleHandler.getStrippedCellStyle2();
        }
    }

    /**
     * Get default cell style
     * @return
     */
    public CellStyle getDefaultCellStyle() {
        return esrStyleHandler.getDefaultCellStyle();
    }

    /**
     * Set new default cell style
     * @param cellStyle New default cell style
     */
    public void setDefaultCellStyle(CellStyle cellStyle) {
        esrStyleHandler.setDefaultCellStyle(cellStyle);
    }

    /**
     * Get green cell style
     * @return
     */
    public CellStyle getGreenCellStyle() {
        return esrStyleHandler.getGreenCellStyle();
    }

    /**
     * Set new green cell style
     * @param cellStyle New default cell style
     */
    public void setGreenCellStyle(CellStyle cellStyle) {
        esrStyleHandler.setGreenCellStyle(cellStyle);
    }

    /**
     * Get yellow cell style
     * @return
     */
    public CellStyle getYellowCellStyle() {
        return esrStyleHandler.getYellowCellStyle();
    }

    /**
     * Set new yellow cell style
     * @param cellStyle New default cell style
     */
    public void setYellowCellStyle(CellStyle cellStyle) {
        esrStyleHandler.setYellowCellStyle(cellStyle);
    }

    /**
     * Get red cell style
     * @return
     */
    public CellStyle getRedCellStyle() {
        return esrStyleHandler.getRedCellStyle();
    }

    /**
     * Set new red cell style
     * @param cellStyle New default cell style
     */
    public void setRedCellStyle(CellStyle cellStyle) {
        esrStyleHandler.setRedCellStyle(cellStyle);
    }

    /**
     * Get header cell style
     * @return
     */
    public CellStyle getHeaderCellStyle() {
        return esrStyleHandler.getHeaderCellStyle();
    }

    /**
     * Set new header cell style
     * @param cellStyle New default cell style
     */
    public void setHeaderCellStyle(CellStyle cellStyle) {
        esrStyleHandler.setHeaderCellStyle(cellStyle);
    }

    /**
     * Get title cell style
     * @return
     */
    public CellStyle getTitleCellStyle() {
        return esrStyleHandler.getTitleCellStyle();
    }

    /**
     * Set new title cell style
     * @param cellStyle New default cell style
     */
    public void setTitleCellStyle(CellStyle cellStyle) {
        esrStyleHandler.setTitleCellStyle(cellStyle);
    }

    /**
     * Set new stripped cell style 1
     * @param cellStyle New default cell style
     */
    public void setStrippedCellStyle1(CellStyle cellStyle) {
        esrStyleHandler.setStrippedCellStyle1(cellStyle);
    }

    /**
     * Set new strinpped cell style 2
     * @param cellStyle New default cell style
     */
    public void setStrippedCellStyle2(CellStyle cellStyle) {
        esrStyleHandler.setStrippedCellStyle2(cellStyle);
    }
}
