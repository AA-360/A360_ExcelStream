package com.automationanywhere.botcommand.samples.commands.utils;

import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.exception.BotCommandException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Array;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class WorkbookHelper {
    public Workbook wb = null;
    private FormulaEvaluator evaluator = null;
    private DataFormatter formatter = new DataFormatter(true);

    public WorkbookHelper(XSSFWorkbook myWorkBook){
        this.wb = myWorkBook;
        this.evaluator = this.wb.getCreationHelper().createFormulaEvaluator();
    }
    public WorkbookHelper(HSSFWorkbook myWorkBook){
        this.wb = myWorkBook;
        this.evaluator = this.wb.getCreationHelper().createFormulaEvaluator();
    }
    public WorkbookHelper(Workbook myWorkBook){
        this.wb = myWorkBook;
        //this.wb.getCreationHelper();
        //this.evaluator = this.wb.getCreationHelper().createFormulaEvaluator();
    }

    public List<String> getSheetsName(){
        List<String> listSheets = new ArrayList();
        Iterator<Sheet> sheetIterator = this.wb.sheetIterator();
        while(sheetIterator.hasNext()){
            Sheet sheet = sheetIterator.next();
            listSheets.add(sheet.getSheetName());
        }

        return listSheets;
    }

    public Boolean sheetExists(String sheetName){
        return this.getSheetsName().contains(sheetName);
    }
    public Boolean sheetExists(Integer sheetIndex){
        return sheetIndex <= (this.getSheetsName().size()-1);
    }
    public StringValue getCellValue(Cell cell,Integer DecimalPoints){
        String value = "";
        DecimalFormat df = new DecimalFormat("#");
        df.setMaximumFractionDigits(DecimalPoints);
        if (cell == null) return new StringValue("");
        try{
            switch (cell.getCellType()) {
                case STRING:
                    value = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    value = df.format(cell.getNumericCellValue());
                    value = value.replace(",",".");
                    break;
                case BOOLEAN:
                    value = Boolean.toString(cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    try {
                        value = df.format(cell.getNumericCellValue());
                        value = value.replace(",",".");
                    }catch(Exception e){
                        value = cell.getStringCellValue();
                    }
                    break;
                case BLANK:
                    value = "";
                    break;
                case ERROR:
                    value = "#REF!";
                    break;
                default:
                    cell.setCellType(CellType.STRING);
                    value = cell.getStringCellValue();
                    value = value==null?"":value;
                    //throw new BotCommandException("ERROR PARSING VALUE TYPE :" + cell.getCellType());
            }
        }catch(Exception e){
            return new StringValue(this.getAsString(cell));

        }
        return new StringValue(value);
    }

    public String getAsString(Cell cl){
        try{
            cl.setCellType(CellType.STRING);
            return cl.getStringCellValue();
        }catch(Exception e){
            return cl.getStringCellValue();
        }
    }

    public List<Row> getRows(Sheet sheet){
        List<Row> listRows = new ArrayList();

        for(Integer rowIdx = 0; rowIdx < sheet.getLastRowNum();rowIdx++){
            listRows.add(sheet.getRow(rowIdx));
        }

//        Iterator<Row> rowIterator = sheet.iterator();
//
//        while (rowIterator.hasNext()) {
//            listRows.add(rowIterator.next());
//        }

        return listRows;
    }

    public List<Cell> getColumns(Row row){
        List<Cell> listCells = new ArrayList();
        Iterator<Cell> cellIterator;
        if(row == null) return listCells;
        //System.out.println(":::::" + row);

        for(Integer colIdx = 0; colIdx < row.getLastCellNum();colIdx++){
            listCells.add(row.getCell(colIdx));
        }

//        row.getLastCellNum();
//
//        try {
//            cellIterator = row.cellIterator();
//        }catch(Exception e){
//            return listCells;
//        }
//
//        while (cellIterator.hasNext()) {
//            listCells.add(cellIterator.next());
//        }
        return listCells;
    }

    public Integer ColumnToIndex(String col){
        String Letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        Integer multiplicador = 1;
        Integer numb = 0;

        char ch;
        for(int i=col.toCharArray().length-1;i>=0;i--) {
            ch = Array.getChar(col.toCharArray(), i);
            numb += multiplicador * (Letters.indexOf(ch)+1);
            multiplicador *=26;
        }
        if(numb > 16384){
            throw new BotCommandException("Column '" + col + "' value invalid!");
        }
        return numb;
    }


    public static XSSFWorkbook createXLSXWorkbook(String file){

        try{
            File myFile = new File(file);
            if(!myFile.exists()){
                throw new BotCommandException("File '" + file + "' not found!");
            }
            FileInputStream fis = new FileInputStream(myFile);

            XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);
            return myWorkBook;
        }catch (IOException e){
            throw new BotCommandException("Error reading/crearing xlsx file:" + e.getMessage());
        }
    }

}
