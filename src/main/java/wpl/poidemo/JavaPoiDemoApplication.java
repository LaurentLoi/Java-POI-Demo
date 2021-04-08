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

        /* If tree nodes arrays in dataTree are not of the same length → works but not as awaited */

        // DATA TREE SHEET 1
        List<List<Integer>> dataTreeSheet01 = new ArrayList<>(); // TREE
        dataTreeSheet01.add(Arrays.asList(1, 2, 4, 5)); // TREE NODE
        dataTreeSheet01.add(Arrays.asList(1, 2, 4, 5));
        dataTreeSheet01.add(Arrays.asList(1, 2, 4, 5));

//        // DATA TREE SHEET 2
//        List<List<Integer>> dataTreeSheet02 = new ArrayList<>();
//        dataTreeSheet02.add(Arrays.asList(1, 3, 2, 2, 1, 2, 2, 3, 2, 4, 3));
//        dataTreeSheet02.add(Arrays.asList(1, 3, 2, 2, 1, 2, 2, 3, 2, 4, 3));
//        dataTreeSheet02.add(Arrays.asList(1, 3, 2, 2, 1, 2, 2, 3, 2, 4, 3));
//
//
//        // DATA TREE SHEET 3
//        List<List<Integer>> dataTreeSheet03 = new ArrayList<>();
//        dataTreeSheet03.add(Arrays.asList(1, 2, 4, 5));
//        dataTreeSheet03.add(Arrays.asList(1, 2, 3, 5));
//        dataTreeSheet03.add(Arrays.asList(1, 2, 4, 5));
//        dataTreeSheet03.add(Arrays.asList(1, 2, 4, 6));

        // LIST OF ALL DATA TREES
        List<List<List<Integer>>> treesList = new ArrayList<>();
        treesList.add(dataTreeSheet01);
//        treesList.add(dataTreeSheet02);
//        treesList.add(dataTreeSheet03);

        /** END DATA **/

        /** VARIABLES GENERATION **/

        // TOTAL SHEETS NUMBER
        int sheetsNumber = treesList.size();

        // TOTAL ROWS COUNTER BY SHEET
        List<Integer> rowsNumbersBySheet = countRowsBySheet(treesList);

        // TOTAL COLS COUNTER BY SHEET
        List<Integer> colsNumbersBySheet = countColsBySheet(treesList);

        // SHEETS TITLES GENERATION
        List<String> sheetsNames = generateSheetsTitles(sheetsNumber);

        // COLS TITLES GENERATION
        List<List<String>> rowsTitles = generateColumnsTitles(sheetsNumber, colsNumbersBySheet);

        /** END VARIABLES GENERATION **/

        /** PROCESS **/
        long startTime = System.currentTimeMillis();

        // FILE INITIALIZATION
        Excel_Init excelInit = new Excel_Init();
        Workbook excelFile = excelInit.excelInitialization(sheetsNumber, sheetsNames);

        // BORDERS + TITLES STYLING
        BordersGen bordersGen = new BordersGen();
        bordersGen.addBorder("bottom", 0, excelFile, colsNumbersBySheet);
        styleGeneratedRows += 1;

        // TITLES FILLING
        TitlesGen titlesGen = new TitlesGen();
        titlesGen.fillTitles(rowsTitles, excelFile);

        // ROWS / COLS GENERATION / CELLS FILLING + INDIVIDUAL CELL STYLING
        CellsGen cellsGen = new CellsGen();
        cellsGen.generateExcelCells(excelFile, colsNumbersBySheet, styleGeneratedRows, treesList);

        // FILE PRODUCTION
        Excel_Producer excelProducer = new Excel_Producer();
        excelProducer.excelFileProduction(excelFile, fileName);

        long totalTime = System.currentTimeMillis() - startTime;
        /** END PROCESS **/

        printStatLogs(sheetsNumber, rowsNumbersBySheet, colsNumbersBySheet, totalTime);
    }

    private static List<Integer> countRowsBySheet(List<List<List<Integer>>> treesList) {
        List<Integer> rowsNumbersBySheet = new ArrayList<>();

        for (List<List<Integer>> dataTreeBySheet : treesList) {
            int totalRowsNumber = 0;
            for (List<Integer> ratioArray : dataTreeBySheet) {
                int currentArrayRowsNumber = 1;
                for (int rowNumber : ratioArray) {
                    currentArrayRowsNumber *= rowNumber;
                }
                totalRowsNumber += currentArrayRowsNumber;
            }
            rowsNumbersBySheet.add(totalRowsNumber);
        }
        return rowsNumbersBySheet;
    }

    private static List<Integer> countColsBySheet(List<List<List<Integer>>> treesList) {
        List<Integer> colsNumbersBySheet = new ArrayList<>();

        for (List<List<Integer>> dataTreesBySheet : treesList) {
            int maxColNumber = 0;
            for (List<Integer> dataTree : dataTreesBySheet) {
                if (dataTree.size() > maxColNumber) {
                    maxColNumber = dataTree.size();
                }
            }
            colsNumbersBySheet.add(maxColNumber);
        }
        return colsNumbersBySheet;
    }

    private static List<String> generateSheetsTitles(int sheetsNumber) {
        List<String> sheetsNames = new ArrayList<>();

        for (int sheetIndex = 0; sheetIndex < sheetsNumber; sheetIndex++) {
            sheetsNames.add("Sheet n°" + (sheetIndex + 1));
        }
        return sheetsNames;
    }

    private static List<List<String>> generateColumnsTitles(int sheetsNumber, List<Integer> colsNumbersBySheet) {
        List<List<String>> rowsTitles = new ArrayList<>();

        for (int sheetIndex = 0; sheetIndex < sheetsNumber; sheetIndex++) {
            List<String> titlesBySheet = new ArrayList<>();

            for (int colIndex = 0; colIndex < colsNumbersBySheet.get(sheetIndex); colIndex++) {
                titlesBySheet.add("Title " + (colIndex + 1) + " - Sheet " + (sheetIndex + 1));
            }
            rowsTitles.add(titlesBySheet);
        }
        return rowsTitles;
    }

    private static int totalColsCounter(List<Integer> colsNumbersBySheet) {
        int totalColNumber = 0;
        for (Integer integer : colsNumbersBySheet) {
            totalColNumber += integer;
        }
        return totalColNumber;
    }

    private static int totalRowsCounter(List<Integer> rowsNumbersBySheet) {
        int totalCellsNumber = 0;
        for (Integer integer : rowsNumbersBySheet) {
            totalCellsNumber += integer;
        }
        return totalCellsNumber;
    }

    private static void printStatLogs(int sheetsNumber, List<Integer> rowsNumbersBySheet, List<Integer> colsNumbersBySheet, long timer) {
        int totalColNumber = totalColsCounter(colsNumbersBySheet);
        int totalRowsNumber = totalRowsCounter(rowsNumbersBySheet);

        System.out.println("Generation of " + sheetsNumber + " sheet(s), " + totalColNumber + " columns and " + totalRowsNumber + " rows is complete.");
        System.out.println("Excel file created. It took " + timer + " milliseconds.");
    }
}
