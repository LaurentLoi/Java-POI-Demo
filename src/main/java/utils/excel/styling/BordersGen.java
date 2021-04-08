package utils.excel.styling;

import org.apache.poi.ss.usermodel.*;

import java.util.List;

public class BordersGen {

    public void addBorder(String direction, int rowToBorder, Workbook currentWorkBook, List<Integer> colsBySheet) {

        // BORDER STYLE INIT
        BorderStyle borderStyle = BorderStyle.THICK;
        // BORDER COLOR INIT
        short borderColor = (short) 0;

        // CELL STYLE INIT
        CellStyle style = currentWorkBook.createCellStyle();
        // VERTICAL / HORIZONTAL ALIGNMENT
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

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
