package utils.excel.fields;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;
@Deprecated
public class RowsGen {

    int sheetIndex = 0;
    // Generate rows
    public Workbook generateExcelRows(Workbook currentWorkBook, List<Integer> rowsNumberBySheet) {

        rowsNumberBySheet.forEach(sheetRows -> {

            Sheet currentSheet = currentWorkBook.getSheetAt(sheetIndex);
            System.out.println("Sheet index : " + sheetIndex);

            for (int i = 0; i < sheetRows; i++) {
                currentSheet.createRow(i);
            }
            sheetIndex += 1;
        });

        sheetIndex = 0;
        return currentWorkBook;
    }

}
