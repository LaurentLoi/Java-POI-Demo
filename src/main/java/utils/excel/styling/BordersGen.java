package utils.excel.styling;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.List;

public class BordersGen {

    public void addBorder(String direction, int rowToBorder, Workbook currentWorkBook, List<Integer> colsBySheet) {

        BorderStyle borderStyle = BorderStyle.THICK;
        short borderColor = (short) 0;
        CellStyle style = currentWorkBook.createCellStyle();

        switch (direction) {
            case "top":
                style.setBorderTop(borderStyle);
                style.setTopBorderColor(borderColor);
                break;

            case "right":
                style.setBorderRight(borderStyle);
                style.setRightBorderColor(borderColor);
                break;

            case "bottom":
                style.setBorderBottom(borderStyle);
                style.setBottomBorderColor(borderColor);
                break;

            case "left":
                style.setBorderLeft(borderStyle);
                style.setLeftBorderColor(borderColor);
                break;
        }

        for (int sheetIndex = 0; sheetIndex < currentWorkBook.getNumberOfSheets(); sheetIndex++) {
            Row row = currentWorkBook.getSheetAt(sheetIndex).createRow(rowToBorder);

            for (int colIndex = 0; colIndex < colsBySheet.get(sheetIndex); colIndex++) {
                row.createCell(colIndex).setCellStyle(style);
            }
        }
    }

}
