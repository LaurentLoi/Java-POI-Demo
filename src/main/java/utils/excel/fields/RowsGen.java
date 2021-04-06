package utils.excel.fields;

import org.apache.poi.ss.usermodel.Sheet;

public class RowsGen {

    // Generate rows
    public Sheet generateExcelRows(Sheet currentSheet, int rowsNumber) {

        for (int i = 0; i <= rowsNumber; i++) {
            currentSheet.createRow(i);
        }
        return currentSheet;
    }

}
