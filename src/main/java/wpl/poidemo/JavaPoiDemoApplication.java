package wpl.poidemo;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import utils.excel.Excel_Init;
import utils.excel.Excel_Producer;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@SpringBootApplication
public class JavaPoiDemoApplication {

    public static void main(String[] args) throws IOException {

        // VARIABLES
        String fileName = "GeneratedExcelFile";
        int sheetsNumber = 3;
        List<String> sheetsNames = Arrays.asList("First sheet", "Second sheet", "Third sheet");

        // FILE INITIALIZATION
        Excel_Init excelInit = new Excel_Init();
        Workbook wb = excelInit.excelInitialization(sheetsNumber, sheetsNames);



        // FILE PRODUCTION
        Excel_Producer excelProducer = new Excel_Producer();
        excelProducer.excelFileProduction(wb, fileName);
    }

}
