package utils.excel.fields;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public class CellsGen {

    // Generate cells
    public Workbook generateExcelCells(Workbook currentWorkBook, List<Integer> rowsNumberBySheet, List<Integer> colsNumberBySheet, int generatedLines) {

        int sheetNumbers = currentWorkBook.getNumberOfSheets();

        for (int sheetIndex = 0; sheetIndex < sheetNumbers; sheetIndex++) {

            Sheet currentSheet = currentWorkBook.getSheetAt(sheetIndex);
            System.out.println("Sheet index : " + sheetIndex);

            for (int rowIndex = generatedLines; rowIndex < (rowsNumberBySheet.get(sheetIndex) + generatedLines); rowIndex++) { // ROWS ITERATION

                Row currentRow = currentSheet.createRow(rowIndex);

                for (int colIndex = 0; colIndex < colsNumberBySheet.get(sheetIndex); colIndex++) { // COL ITERATION

                    Cell cell = currentRow.createCell(colIndex);
                    cell.setCellValue(sheetIndex + "." + colIndex + "." + (currentRow.getRowNum() - generatedLines));

                }
            }
        };

        return currentWorkBook;
    }

}
