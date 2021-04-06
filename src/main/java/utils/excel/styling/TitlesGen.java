package utils.excel.styling;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public class TitlesGen {

    public void fillTitles(List<List<String>> rowsTitles, Workbook currentWorkBook) {

        for (int sheetIndex = 0; sheetIndex < rowsTitles.size(); sheetIndex++) {

            Sheet currentSheet = currentWorkBook.getSheetAt(sheetIndex);
            List<String> sheetTitles = rowsTitles.get(sheetIndex);

            for (int colIndex = 0; colIndex < sheetTitles.size(); colIndex++) {
                currentSheet.getRow(0).getCell(colIndex).setCellValue(sheetTitles.get(colIndex));
                currentWorkBook.getSheetAt(sheetIndex).autoSizeColumn(colIndex); // AUTO WIDTH
            }

        }

    }

}
