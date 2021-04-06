package utils.excel;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

public class Excel_Producer {

    public void excelFileProduction(Workbook excelFile, String fileName) throws IOException {
        OutputStream fileOut = new FileOutputStream(fileName + ".xslx");
        excelFile.write(fileOut);
        excelFile.close();
    }

}
