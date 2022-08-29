package com.automationanywhere.botcommand.samples.commands.basic;

import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.data.impl.TableValue;
import com.automationanywhere.botcommand.data.model.table.Table;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.samples.commands.utils.FindInListSchema;
import com.automationanywhere.botcommand.samples.commands.utils.WorkbookHelper;
import com.automationanywhere.commandsdk.annotations.*;
import com.automationanywhere.commandsdk.annotations.rules.FileExtension;
import com.automationanywhere.commandsdk.annotations.rules.GreaterThanEqualTo;
import com.automationanywhere.commandsdk.annotations.rules.NotEmpty;
import com.automationanywhere.commandsdk.model.DataType;
import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import static com.automationanywhere.commandsdk.model.AttributeType.*;

//BotCommand makes a class eligible for being considered as an action.
@BotCommand

//CommandPks adds required information to be displayable on GUI.
@CommandPkg(
        //Unique name inside a package and label to display.
        name = "SXlsToTable",
        label = "SXlsToTable",
        node_label = "{{file}} to DataTable",
        description = "",
        icon = "pkg.svg",
        return_type = DataType.TABLE,
        return_required = true,
        return_description = "DataTable from xls file"
)

public class SXlsToTable {
    @Execute
    public TableValue action(
            @Idx(index = "1", type = FILE)
            @Pkg(label = "XLS file",description = "example: C:\\folder\\file.xls")
            @FileExtension("xls")
            @NotEmpty
            String file,

            @Idx(index = "2", type = SELECT, options = {
                    @Idx.Option(index = "2.1", pkg = @Pkg(label = "ByName", value = "name")),
                    @Idx.Option(index = "2.2", pkg = @Pkg(label = "ByIndex", value = "index"))})
            @Pkg(label = "Sheet:", description = "", default_value = "name", default_value_type = DataType.STRING)
            @NotEmpty
            String getSheetBy,

            @Idx(index = "2.1.1", type = TEXT)
            @Pkg(label = "Insira o nome da sheet:",description = "Será adiconada no fim da tabela")
            @NotEmpty
            String sheetName,

            @Idx(index = "2.2.1", type = NUMBER)
            @Pkg(label = "Insira o index da sheet:")
            @NotEmpty
            @GreaterThanEqualTo(value = "0")
            Double sheetIndex,

            @Idx(index = "3", type = TEXT)
            @Pkg(label = "Insira as colunas desejadas:", description = "Exemplo: A:C ou A|B|C ")
            @NotEmpty
            String Columns,

            @Idx(index = "4", type = CHECKBOX)
            @Pkg(label = "Contem cabeçalhos",default_value = "false",default_value_type = DataType.BOOLEAN)
            @NotEmpty
            Boolean hasHeaders,

            @Idx(index = "5", type = CHECKBOX)
            @Pkg(label = "Linha de início",default_value_type = DataType.BOOLEAN,default_value = "false")
            @NotEmpty
                    Boolean RowStartCheck,

            @Idx(index = "5.1", type = NUMBER)
            @Pkg(label = "Insira o numero da linha:",description = "tem que ser maior que 0",default_value_type = DataType.NUMBER,default_value = "1")
            @NotEmpty
            @GreaterThanEqualTo(value = "1")
                    Double RowStart,
            @Idx(index = "6", type = CHECKBOX)
            @Pkg(label = "Casas Decimais",default_value = "false",default_value_type = DataType.BOOLEAN)
            @NotEmpty
                    Boolean Decimal,
            @Idx(index = "6.1", type = NUMBER)
            @Pkg(label = "Pontos Decimais:")
            @NotEmpty
            @GreaterThanEqualTo(value = "0")
                    Double DecimalPoints

) {

        //=============================================================================== VALIDATE
        if ("".equals(file.trim()))
            throw new BotCommandException("Please select a valid file for processing.");

        if(!file.toUpperCase().endsWith(".XLS")){
            throw new BotCommandException("Please select a supported file to continue");
        }

        try (
                InputStream is = new FileInputStream(new File(file));
                Workbook myWorkBook = new HSSFWorkbook(is);
        ){
        //================================================================= CREATE WORKBOOK OBJECT
            WorkbookHelper wbH = new WorkbookHelper(myWorkBook);
            //================================================================= VALIDATE RANGE COLUMNS
            List<Integer> colsToreturn = this.columnsToReturn(Columns,wbH);

            //================================================================= GET SHEET
            Sheet mySheet = this.getSheet(getSheetBy,sheetName,sheetIndex,wbH);

            //================================================================= GET ROWS
            List<com.automationanywhere.botcommand.data.model.table.Row> listRows= new ArrayList<>();
            List<String> HEADERS = new ArrayList<>();
            DecimalPoints = Decimal?DecimalPoints:0.0;

            Integer lastRowNum =0;

            for(Row rw : mySheet){

                List<Value> rwValue = new ArrayList<>();

                //System.out.println("ROW:"+ rw.getRowNum() + " LIST:"+ listRows.size()+1);
                if((rw.getRowNum() > 0 && rw.getRowNum() == listRows.size()) || rw.getRowNum() == 0){

                }else{
                    Integer dif = rw.getRowNum() - lastRowNum;
                    for(int i=0;i<dif;i++) {
                        rwValue = new ArrayList<>();
                        for (Integer colIdx : colsToreturn) {
                            rwValue.add(new StringValue(""));
                        }
                        listRows.add(new com.automationanywhere.botcommand.data.model.table.Row(rwValue));
                    }
                }

                rwValue = new ArrayList<>();
                List<Cell> listCol = wbH.getColumns(rw);
                for(Integer colIdx: colsToreturn){
                    if(colIdx <= (listCol.size()-1)){
                        Cell col = listCol.get(colIdx);
                        rwValue.add(wbH.getCellValue(col,DecimalPoints.intValue()));
                    }else{
                        rwValue.add(new StringValue(""));
                    }
                }
                listRows.add(new com.automationanywhere.botcommand.data.model.table.Row(rwValue));

                lastRowNum = rw.getRowNum();
            }


            //============================================================================================ NOME DAS COLUNAS
            Integer idx = 0;
            RowStart = RowStart ==null?1:RowStart;
            for(int i=0;i<colsToreturn.size();i++) {
                if (hasHeaders) {
                    HEADERS.add(listRows.get(RowStart.intValue()-1).getValues().get(i).toString());
                } else {
                    HEADERS.add(idx.toString());
                }
                idx++;
            }
            //========================================================================================== DELETANDO LINHAS
            if(RowStartCheck){
                for(int i=0;i<RowStart-(hasHeaders?0:1);i++) {
                    listRows.remove(0);
                }
            }else{
                listRows.remove(0);
            }

            FindInListSchema fnd = new FindInListSchema(HEADERS);

            System.out.println(HEADERS);
            Table OUTPUT = new Table(fnd.schemas,listRows);

            return new TableValue(OUTPUT);
        }catch(IOException e){
            throw new BotCommandException("Error: " + e.getMessage());
        }
    }

        private XSSFWorkbook createXLSXWorkbook(String file){

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

        private List<Integer> columnsToReturn(String Columns, WorkbookHelper wbH){
                List<Integer> colsIndex = new ArrayList<>();
                Columns = Columns.toUpperCase().trim();
                Boolean pattern1= Columns.matches("^([A-Z]{1,3}):([A-Z]{1,3})$");
                Boolean pattern2= Columns.matches("^(([A-Z]{1,3})\\|)*[A-Z]{1,3}$");

                if(!(pattern1 || pattern2)){
                        throw new BotCommandException("Columns (" + Columns + ") has not a valid format try to use as A:C or A|B|C");
                }
                if(pattern1){
                        String[] addrs = Columns.split(":");
                        colsIndex = this.getNumbersInRange(wbH.ColumnToIndex(addrs[0])-1,wbH.ColumnToIndex(addrs[1])-1);
                }else{
                        String[] addrs = Columns.split("\\|");
                        for(String cel: addrs){
                                colsIndex.add(wbH.ColumnToIndex(cel)-1);
                        }
                }
                return colsIndex;
        }

        private Sheet getSheet(String getSheetBy, String sheetName, Double sheetIndex, WorkbookHelper wbH){
                Sheet mySheet = null;
                if(getSheetBy.equals("name")){
                        if(wbH.sheetExists(sheetName)){
                                wbH.wb.getSheet(sheetName);
                                mySheet = wbH.wb.getSheet(sheetName);
                        }else{
                                throw new BotCommandException("Sheet '" + sheetName + "' not found!");
                        }
                }else {
                        if(wbH.sheetExists(sheetIndex.intValue())){
                                mySheet = wbH.wb.getSheetAt(sheetIndex.intValue());
                        }else{
                                throw new BotCommandException("Sheet index '" + sheetIndex.intValue() + "' not found!");
                        }
                }
                return mySheet;
        }

        public List<Integer> getNumbersInRange(int start, int end) {
                List<Integer> result = new ArrayList<>();
                for (int i = start; i <= end; i++) {
                        result.add(i);
                }
                return result;
        }

}
