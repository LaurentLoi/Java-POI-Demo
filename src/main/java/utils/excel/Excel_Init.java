package utils.excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;

public class Excel_Init {

    public Workbook excelInitialization(int sheetsNumber, String[] sheetsNames) throws FileNotFoundException {

        // Init file
        Workbook excelFile = new XSSFWorkbook();

        // init sheets
        for (int i = 0; i < sheetsNumber; i++) {
            excelFile.createSheet(sheetsNames[i]);
        }

        return excelFile;

    }
}
