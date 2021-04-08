package wpl.poidemo;

import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import utils.excel.Excel_Init;
import utils.excel.Excel_Producer;
import utils.excel.fields.CellsGen;
import utils.excel.styling.BordersGen;
import utils.excel.styling.TitlesGen;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

@SpringBootApplication
public class JavaPoiDemoApplication {

    public static void main(String[] args) throws IOException {

        // AUTO GENERATED ROWS COUNTER (TITLE, ...)
        int styleGeneratedRows = 0;

        /** DATA TO INSERT **/
        // FILE NAME
        String fileName = "GeneratedExcelFile";

        // RATIO TABLES SHEET 1
        List<List<Integer>> ratioSheet1 = new ArrayList<>();
        ratioSheet1.add(Arrays.asList(1, 2, 4, 5));
        ratioSheet1.add(Arrays.asList(1, 2, 4, 5));
        ratioSheet1.add(Arrays.asList(1, 2, 4, 5));

//        // RATIO TABLES SHEET 2
//        List<List<Integer>> ratioSheet2 = new ArrayList<>();
//        ratioSheet2.add(Arrays.asList(1, 3, 2, 2, 1, 2, 2, 3, 2, 4, 3));
//        ratioSheet2.add(Arrays.asList(1, 3, 2, 2, 1, 2, 2, 3, 2, 4, 3));
//        ratioSheet2.add(Arrays.asList(1, 3, 2, 2, 1, 2, 2, 3, 2, 4, 3));


//        // RATIO TABLES SHEET 3
//        List<List<Integer>> ratioSheet3 = new ArrayList<>();
//        ratioSheet3.add(Arrays.asList(1, 2, 4, 5));
//        ratioSheet3.add(Arrays.asList(1, 2, 4, 5));
//        ratioSheet3.add(Arrays.asList(1, 2, 4, 5));
//        ratioSheet3.add(Arrays.asList(1, 2, 4, 5));

        List<List<List<Integer>>> ratioChart = new ArrayList<>(); //Arrays.asList(ratioSheet1, ratioSheet2, ratioSheet3);
        ratioChart.add(ratioSheet1);
//        ratioChart.add(ratioSheet2);
//        ratioChart.add(ratioSheet3);

        /** END DATA **/

        /** VARIABLES GENERATION **/

        // TOTAL ROWS COUNTER BY SHEET
        List<Integer> rowsNumbersBySheet = new ArrayList<>();
        for (List<List<Integer>> ratioListBySheet : ratioChart) {
            int totalRowsNumber = 0;
            for (List<Integer> ratioArray : ratioListBySheet) {
                int currentArrayRowsNumber = 1;
                for (int rowNumber : ratioArray) {
                    currentArrayRowsNumber *= rowNumber;
                }
                totalRowsNumber += currentArrayRowsNumber;
            }
            rowsNumbersBySheet.add(totalRowsNumber);
        }

        // TOTAL COLS COUNTER  BY SHEET
        List<Integer> colsNumbersBySheet = new ArrayList<>();
        for (List<List<Integer>> ratioListBySheet : ratioChart) {
            int colNumber = 0;
            for (List<Integer> ratioArray : ratioListBySheet) {
                colNumber = ratioArray.size();
            }
            colsNumbersBySheet.add(colNumber);
        }

        // SHEETS TITLES GENERATION
        int sheetsNumber = ratioChart.size(); // number of sheet by file
        List<String> sheetsNames = new ArrayList<>();

        for (int sheetIndex = 0; sheetIndex < sheetsNumber; sheetIndex++) {
            sheetsNames.add("Sheet n°" + (sheetIndex + 1));
        }

        // COLS TITLES GENERATION
        List<List<String>> rowsTitles = new ArrayList<>();
        for (int sheetIndex = 0; sheetIndex < sheetsNumber; sheetIndex++) {

            List<String> titlesBySheet = new ArrayList<>();

            for (int colIndex = 0; colIndex < colsNumbersBySheet.get(sheetIndex); colIndex++) {
                titlesBySheet.add("Title " + (colIndex + 1) + " - Sheet " + (sheetIndex + 1));
            }

            rowsTitles.add(titlesBySheet);
        }

        /** END VARIABLES **/

        /** PROCESS **/
        long startTime = System.currentTimeMillis();

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

        // ROWS / COLS GENERATION / CELLS FILLING + INDIVIDUAL CELL STYLING
        CellsGen cellsGen = new CellsGen();
        cellsGen.generateExcelCells(excelFile, colsNumbersBySheet, styleGeneratedRows, ratioChart);

        // FILE PRODUCTION
        Excel_Producer excelProducer = new Excel_Producer();
        excelProducer.excelFileProduction(excelFile, fileName);

        long endTime = System.currentTimeMillis();
        /** END PROCESS **/

        // COL COUNTER
        int totalColNumber = 0;
        for (Integer integer : colsNumbersBySheet) {
            totalColNumber += integer;
        }

        // ROW COUNTER
        int totalCellsNumber = 0;
        for (Integer integer : rowsNumbersBySheet) {
            totalCellsNumber += integer;
        }

        System.out.println("\n");
        System.out.println("Generation of " + sheetsNumber + " sheet(s), " + totalColNumber + " columns and " + totalCellsNumber + " rows is complete.");
        System.out.println("Excel file created. It took " + (endTime - startTime) + " milliseconds.");
    }

}
