# ESRConstructor6

## General information

A framework that simplifies the creation of simple reports using Apache POI. 
Library is compiled for Java 1.6 and uses Apache POI 3.17.

## Report entities
General report outline:
<img width="559" height="525" alt="image" src="https://github.com/user-attachments/assets/d5d27f95-58c3-4424-901a-6c67f72c8e39" />


| Object  | Description | Objects that can be added |
| ------------- | ------------- | ------------- |
| ESRConstructor  | The main object to which the HSSFWorkBook object must be passed.  | ESRReport (multiple)  |
| ESRReport  | Report object. May consist of descriptive information, a header block, and a main table composed of blocks.  | ESRReport (multiple)<br>ESRHeaderBlock<br>ESRTableBlock (multiple)<br>ESRSingleRowBlock (multiple)  |
| ESRTitleBlock  | Descriptive information block. It combines the maximum number of columns available in the report. | ESRTitleRow (multiple)  |
| ESRTitleRow  | DRow of descriptive information section. |   |
| ESRHeaderBlock  | Header block. Repeats on every new page of the report. | ESRHeaderRow (multiple)<br>ESRMerge |
| ESRHeaderRow  | Row of header block. | ESRHeaderCell (multiple) |
| ESRHeaderCell  | Cell of header section row. | |
| ESRTableBlock  | Table block of main information. Row numbering in each block starts at 0, which makes it easier to define merger regions. | ESRTableRow (multiple)<br>ESRMerge |
| ESRSingleRowBlock  | A block of main information consisting of a single row. This makes it easier to create reports that do not require row merging. Please note that adding a row will always overwrite the only row currently in the block that may already exist. | ESRTableRow |
| ESRTableRow  | Row of main information block. | ESRTableCell (multiple) |
| ESRTableCell  | Cell of main information block row. | |
| ESRMerge  | An object containing a list of cell merges within a block. Note that row numbering in each block starts at 0. | |

## General procedure for creating a report
1. Create a new ESRConstructor6 report builder and pass the created Workbook to it. There can be used HSSFWorkbook, XSSFWorkbook, SXSSFWorkbook.
2. Create an ESRReport. The report may contain a preliminary information block (ESRTitleBlock), a header block (ESRHeaderBlock), and multiple information blocks (ESRTableBlock) that hold the report data. Note that you can combine multiple reports within a single constructor. Each report will start on a new sheet relative to the previous one.
3. The ESRTitleBlock is used to add information at the beginning of a report (for example, a row describing the report and other introductory information). This block can contain multiple rows. It will always be positioned at the very beginning of the report. It will always span the maximum number of columns used in the report. The block can contain multiple text rows, which are added using the addTitleRow(String titleRow) method. You can apply the CellStyle style to the block.

## Examples
[Examples EN](EXAMPLES_EN.md)
