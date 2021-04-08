package utils.excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.util.List;

public class Excel_Init {

    public Workbook excelInitialization(int sheetsNumber, List<String> sheetsNames) {

        // INIT FILE
        Workbook excelFile = new XSSFWorkbook();

        // INIT SHEETS
        for (int i = 0; i < sheetsNumber; i++) {
            excelFile.createSheet(sheetsNames.get(i));
        }

        return excelFile;
    }
}
