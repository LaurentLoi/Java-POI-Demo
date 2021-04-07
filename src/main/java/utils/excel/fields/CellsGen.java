package utils.excel.fields;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

public class CellsGen {

    // Generate cells
    public Workbook generateExcelCells(Workbook currentWorkBook, List<Integer> rowsNumberBySheet, List<Integer> colsNumberBySheet, int generatedLines, List<List<List<Integer>>> ratioChart) {

        CellStyle cellStyle = currentWorkBook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        int sheetNumbers = currentWorkBook.getNumberOfSheets();

        // SHEETS ITERATION
        for (int sheetIndex = 0; sheetIndex < sheetNumbers; sheetIndex++) {
            int maxColNumber = colsNumberBySheet.get(sheetIndex);
            Sheet currentSheet = currentWorkBook.getSheetAt(sheetIndex);
            List<List<Integer>> currentRatioList = ratioChart.get(sheetIndex);


            List<Integer> totalCellCounter = new ArrayList<>();
            for (int i = 0; i < maxColNumber; i++) {
                totalCellCounter.add(i, 0);
            }

            //RATIO LIST ITERATOR
            for (int ratioListIndex = 0; ratioListIndex < currentRatioList.size(); ratioListIndex++) {
                int cellCounter = 0;
                List<Integer> currentRatio = currentRatioList.get(ratioListIndex);
                System.out.println("Currently working on sheet " + (sheetIndex + 1) + " with ratio list : " + (ratioListIndex + 1));
                // COL ITERATOR
                for (int colIndex = 0; colIndex < maxColNumber; colIndex++) {

                    System.out.println("With COL : " + colIndex);

                    int cellsNbrToInsert = cellsToInsertCounter(currentRatio, colIndex);
                    int currentTotalCellCounter = totalCellCounter.get(colIndex);

                    long startTime = System.currentTimeMillis();
                    for (int i = 0; i < cellsNbrToInsert; i++) {

                        int cellIndex = cellCounter + generatedLines + currentTotalCellCounter;
                        Row currentRow = currentSheet.getRow(cellIndex);
                        Cell currentCell;

                        // if first iteration → create row
                        if (currentRow == null) {
                            currentRow = currentSheet.createRow(cellIndex);
                            currentCell = currentRow.createCell(colIndex);
                        }

                        currentCell = currentRow.createCell(colIndex);
                        currentCell.setCellValue((sheetIndex) + "." + colIndex + "." + i + ".");

                        // sets individual cell style
                        currentCell.setCellStyle(cellStyle);

                        if (colIndex != (maxColNumber - 1)) {
                            int mergeEndIndex = this.stackSplitter(currentRatio, colIndex);

                            currentSheet.addMergedRegion(new CellRangeAddress(
                                    cellIndex, //first row (0-based)
                                    (cellIndex + mergeEndIndex) - 1, //last row  (0-based)
                                    colIndex, //first column (0-based)
                                    colIndex  //last column  (0-based)
                            ));
                            cellCounter += mergeEndIndex;
                        } else {
                            cellCounter += 1;
                        }
                    }
                    totalCellCounter.set(colIndex, totalCellCounter.get(colIndex) + cellCounter);
                    cellCounter = 0;
                    long endTime = System.currentTimeMillis();
                    System.out.println(cellsNbrToInsert + " cells inserted in " + (endTime - startTime) + " milliseconds");
                }
            }
            // sets default height
            currentSheet.setDefaultRowHeight((short) 420);
        }
        return currentWorkBook;
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
