package utils.excel.fields;

import org.apache.poi.ss.usermodel.Sheet;

public class CellsGen {

    // Generate cells
    private Sheet generateExcelCells(Sheet currentSheet, int colsNumber) {

        currentSheet.rowIterator().forEachRemaining(row -> {
            for (int i = 0; i < colsNumber; i++) {
                row.createCell(i);
            }
        });

        return currentSheet;
    }

}
