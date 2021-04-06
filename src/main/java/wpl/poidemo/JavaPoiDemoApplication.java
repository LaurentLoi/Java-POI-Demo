package wpl.poidemo;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import utils.excel.Excel_Init;
import utils.excel.Excel_Producer;
import utils.excel.fields.CellsGen;
import utils.excel.fields.RowsGen;
import utils.excel.styling.BordersGen;
import utils.excel.styling.TitlesGen;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@SpringBootApplication
public class JavaPoiDemoApplication {

    public static void main(String[] args) throws IOException {

        int styleGeneratedRows = 0;

        /** VARIABLES **/

        // FILE
        String fileName = "GeneratedExcelFile";

        // SHEETS
        int sheetsNumber = 3; // number of sheet by file
        List<String> sheetsNames = Arrays.asList("First sheet", "Second sheet", "Third sheet");

        // RATIO TABLES SHEET 1
        List<List<Integer>> ratioSheet1 = new ArrayList<>();
        ratioSheet1.add(Arrays.asList(1, 3, 10, 5, 100));// 15 000 lines
        ratioSheet1.add(Arrays.asList(1, 4, 12, 7, 48)); // 16 128 lines
        ratioSheet1.add(Arrays.asList(1, 6, 8, 12, 20)); // 11 520 lines
        //                   TOTAL = 3, 13, 30, 24, 168 -   42 648 lines

        // RATIO TABLES SHEET 2
        List<List<Integer>> ratioSheet2 = new ArrayList<>();
        ratioSheet1.add(Arrays.asList(1, 8, 9, 35));  // 2 520 lines
        ratioSheet1.add(Arrays.asList(1, 6, 15, 80)); // 7 200 lines
        //                   TOTAL = 2, 14, 24, 115 -    9 720 lines

        // RATIO TABLES SHEET 3
        List<List<Integer>> ratioSheet3 = new ArrayList<>();
        ratioSheet1.add(Arrays.asList(1, 3, 9, 12));  // 324 lines
        ratioSheet1.add(Arrays.asList(1, 6, 6, 25));  // 900 lines
        ratioSheet1.add(Arrays.asList(1, 7, 4, 17));  // 476 lines
        ratioSheet1.add(Arrays.asList(1, 5, 8, 30)); // 1200 lines
        //                   TOTAL = 4, 21, 27, 84 -   2 900 lines

        // ROWS
        List<Integer> rowsNumbersBySheet = new ArrayList<>(); // number of rows by sheet
        rowsNumbersBySheet.add(42648); // SHEET 1
        rowsNumbersBySheet.add(9720); // SHEET 2
        rowsNumbersBySheet.add(2900); // SHEET 3

        // COLS
        List<Integer> colsNumbersBySheet = new ArrayList<>();
        colsNumbersBySheet.add(5); // SHEET 1
        colsNumbersBySheet.add(4); // SHEET 2
        colsNumbersBySheet.add(4); // SHEET 3

        // COLS TITLES
        List<List<String>> rowsTitles = new ArrayList<>();
        rowsTitles.add(Arrays.asList("Title 1 - sheet 1", "Title 2 - sheet 1", "Title 3 - sheet 1", "Title 4 - sheet 1", "Title 5 - sheet 1"));
        rowsTitles.add(Arrays.asList("Title 1 - sheet 2", "Title 2 - sheet 2", "Title 3 - sheet 2", "Title 4 - sheet 2"));
        rowsTitles.add(Arrays.asList("Title 1 - sheet 3", "Title 2 - sheet 3", "Title 3 - sheet 3", "Title 4 - sheet 3"));

        /** END VARIABLES **/


        long startTime = System.currentTimeMillis();

        /** PROCESS **/

        // FILE INITIALIZATION
        Excel_Init excelInit = new Excel_Init();
        Workbook excelFile = excelInit.excelInitialization(sheetsNumber, sheetsNames);

        // BORDERS  + TITLES STYLING
        BordersGen bordersGen = new BordersGen();
        bordersGen.addBorder("bottom", 0, excelFile, colsNumbersBySheet);
        styleGeneratedRows += 1;

        // TITLES FILLING
        TitlesGen titlesGen = new TitlesGen();
        titlesGen.fillTitles(rowsTitles, excelFile);

        // ROWS / COLS GENERATION & CELLS FILLING + INDIVIDUAL CELL STYLING
        CellsGen cellsGen = new CellsGen();
        cellsGen.generateExcelCells(excelFile, rowsNumbersBySheet, colsNumbersBySheet, styleGeneratedRows);

        /** TODO
        *  - Generate borders
         *      - for
        *  - Generate titles
        *  - Generate rows
        *  - Generate cells
        *  - Fill data
        *  - Generate style
        ********************/


        // FILE PRODUCTION
        Excel_Producer excelProducer = new Excel_Producer();
        excelProducer.excelFileProduction(excelFile, fileName);

        /** END PROCESS **/

        long endTime = System.currentTimeMillis();
        System.out.println("Excel file generated. It took " + (endTime - startTime) + " milliseconds.");
    }

}
