package utils.excel.fields;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public class CellsGen {

    // Generate cells
    public Workbook generateExcelCells(Workbook currentWorkBook, List<Integer> rowsNumberBySheet, List<Integer> colsNumberBySheet) {

        int sheetNumbers = currentWorkBook.getNumberOfSheets();

        for (int sheetIndex = 0; sheetIndex < sheetNumbers; sheetIndex++) {

            Sheet currentSheet = currentWorkBook.getSheetAt(sheetIndex);
            System.out.println("Sheet index : " + sheetIndex);

            for (int rowIndex = 0; rowIndex < rowsNumberBySheet.get(sheetIndex); rowIndex++) {

                Row currentRow = currentSheet.createRow(rowIndex);

                for (int colIndex = 0; colIndex < colsNumberBySheet.get(sheetIndex); colIndex++) {

                    Cell cell = currentRow.createCell(colIndex);
                    cell.setCellValue(sheetIndex + "." + colIndex + "." + currentRow.getRowNum());

                }
            }
        };

        return currentWorkBook;
    }

}
