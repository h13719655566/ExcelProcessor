package com.rlung;

import org.apache.poi.ss.usermodel.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Set;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Main {
    public static void main(String[] args) throws IOException {
        //Data input
        String inputFilePath = "your path";
        String outputFilePath = "output path";

        Set<String> values = new HashSet<>();

        FileInputStream inputStream = new FileInputStream(inputFilePath);
        Workbook mergeWorkbook = WorkbookFactory.create(true);
        Workbook inputWorkbook = WorkbookFactory.create(inputStream);
        int numSheets = 0;
        numSheets = inputWorkbook.getNumberOfSheets();
        Boolean isBlank = false;
        Boolean firstRound = true;

        for(int i = 0; i<numSheets; i++){
            Sheet sheet = inputWorkbook.getSheetAt(i);
            Sheet newSheet = mergeWorkbook.createSheet(sheet.getSheetName());

            int rowNum = 0;
            for(Row row : sheet){
                if(!isBlank&&!firstRound){
                    rowNum++;
                }
                firstRound = false;
                //don't create new row if it's blank
                Row newRow = newSheet.createRow(rowNum);
                isBlank = true;

                int colNum = 0;
                int x=0;
                for(Cell cell : row){
                    //Print
                    String tmp = cell.getStringCellValue();
                    String value = tmp.toUpperCase();
                    if(x==1){
                        x=0;
                        continue;
                    }
                    //delete the duplicate data
                    if (values.add(value)) {
                        Cell newcell = newRow.createCell(colNum++);
                        newcell.setCellValue(tmp);
                        System.out.println(tmp);
                        System.out.println(rowNum);
                        isBlank = false;
                    }else {
                        x=1;
                    }
                }
            }
        }
        //output DATA
        FileOutputStream fileout = new FileOutputStream(outputFilePath);
        mergeWorkbook.write(fileout);
        fileout.close();
    }
}