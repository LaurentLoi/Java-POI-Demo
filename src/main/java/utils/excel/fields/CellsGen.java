package utils.excel.fields;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

public class CellsGen {

    // GENERATE CELLS
    public void generateExcelCells(Workbook currentWorkBook, List<Integer> colsNumberBySheet, int generatedLines, List<List<List<Integer>>> treesList) {

        // CELL STYLE INIT
        CellStyle cellStyle = currentWorkBook.createCellStyle();

        // VERTICAL / HORIZONTAL ALIGNMENT
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        int sheetNumbers = currentWorkBook.getNumberOfSheets();

        // SHEETS ITERATION
        for (int sheetIndex = 0; sheetIndex < sheetNumbers; sheetIndex++) {
            int maxColNumber = colsNumberBySheet.get(sheetIndex);
            Sheet currentSheet = currentWorkBook.getSheetAt(sheetIndex);
            List<List<Integer>> currentDataTree = treesList.get(sheetIndex);


            List<Integer> totalCellCounter = new ArrayList<>();
            for (int i = 0; i < maxColNumber; i++) {
                totalCellCounter.add(i, 0);
            }

            //RATIO LIST ITERATOR
            for (List<Integer> treeNode : currentDataTree) {
                int cellCounter = 0;

                // COL ITERATOR
                for (int colIndex = 0; colIndex < maxColNumber; colIndex++) {

                    int cellsNbrToInsert = cellsToInsertCounter(treeNode, colIndex);
                    int currentTotalCellCounter = totalCellCounter.get(colIndex);

                    for (int i = 0; i < cellsNbrToInsert; i++) {

                        int cellIndex = cellCounter + generatedLines + currentTotalCellCounter;
                        Row currentRow = currentSheet.getRow(cellIndex);
                        Cell currentCell;

                        // IF FIRST ITERATION → CREATE ROW
                        if (currentRow == null) {
                            currentRow = currentSheet.createRow(cellIndex);
                        }

                        // CREATE CELL
                        currentCell = currentRow.createCell(colIndex);
                        // FILL CELL
                        currentCell.setCellValue((sheetIndex) + "." + colIndex + "." + i + ".");
                        // SETS INDIVIDUAL CELL STYLE
                        currentCell.setCellStyle(cellStyle);

                        // IF NOT LAST COL → MERGE CELLS
                        if (colIndex != (maxColNumber - 1)) {
                            int mergeEndIndex = this.stackSplitter(treeNode, colIndex);

                            currentSheet.addMergedRegion(new CellRangeAddress(
                                    cellIndex,
                                    (cellIndex + mergeEndIndex) - 1,
                                    colIndex,
                                    colIndex
                            ));
                            cellCounter += mergeEndIndex;
                        } else {
                            cellCounter += 1;
                        }
                    }
                    totalCellCounter.set(colIndex, totalCellCounter.get(colIndex) + cellCounter);
                    cellCounter = 0;
                }
            }
            // SETS SHEET DEFAULT ROW HEIGHT
            currentSheet.setDefaultRowHeight((short) 420);
        }
    }

    private int cellsToInsertCounter(List<Integer> currentRatio, int colIndex) {
        int counter = 1;
        for (int i = 0; i <= colIndex; i++) {
            counter *= currentRatio.get(i);
        }
        return counter;
    }

    private int stackSplitter(List<Integer> dataStack, int colIndex) {
        int currentStackSize = 1;
        for (int i = colIndex + 1; i < dataStack.size(); i++) {
            currentStackSize *= dataStack.get(i);
        }
        return currentStackSize;
    }
}
