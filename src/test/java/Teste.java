import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.impl.StringValue;
import com.automationanywhere.botcommand.data.impl.TableValue;
import com.automationanywhere.botcommand.data.model.table.Table;
import com.automationanywhere.botcommand.exception.BotCommandException;
import com.automationanywhere.botcommand.samples.commands.basic.SXlsToTable;
import com.automationanywhere.botcommand.samples.commands.basic.SXlsxToTable;
import com.automationanywhere.botcommand.samples.commands.utils.FindInListSchema;
import com.automationanywhere.botcommand.samples.commands.utils.WorkbookHelper;
import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import java.util.ArrayList;
import java.util.List;


public class Teste {

    @Test
    public void tt() {
        //SXlsToTable a = new SXlsToTable();
        SXlsxToTable a = new SXlsxToTable();

        //String File = "C:\\Users\\melque\\Documents\\teste3.xls";
        String File = "C:\\Users\\melque\\Documents\\Input_correto.xlsx";

        //========================================= GET SHEET
        String getSheetBy = "index";// "index"
        String sheetName = "PROCSPCA";
        Double sheetIndex = 0.0;
        Boolean RowStartCheck = false;
        Double RowStart = null;
        //======================================== GET RANGE

        //======================================== OTHERS
        Boolean hasHeaders = true;
        String Columns = "A:I";

        TableValue tbs = a.action(File,getSheetBy,sheetName,sheetIndex,Columns,hasHeaders,RowStartCheck,RowStart,false,4.0);
        uteisTest.printTable(tbs.get(),15);
        System.out.println(tbs.get().getRows().size());

    }


    //@Test
    public void t(){
        try (
                InputStream is = new FileInputStream(new File("C:\\Users\\melque\\Documents\\teste3.xlsx"));
                Workbook workbook = StreamingReader.builder()
                        .rowCacheSize(1000)
                        .bufferSize(4096)
                        .open(is);
        ) {

            WorkbookHelper wbH = new WorkbookHelper(workbook);
            //Sheet mySheet = workbook.getSheet("AGNIVJU");


//            for (Row rw : mySheet) {
//                System.out.println(rw);
//            }

            //================================================================= CREATE WORKBOOK OBJECT
            String Columns = "A:V";
            String getSheetBy = "name";
            String sheetName = "AGNIVJU";
            Double sheetIndex = 0.0;
            Boolean hasHeaders = false;
            //================================================================= VALIDATE RANGE COLUMNS
            List<Integer> colsToreturn = this.columnsToReturn(Columns, wbH);

            //================================================================= GET SHEET
            Sheet mySheet = this.getSheet(getSheetBy, sheetName, sheetIndex, wbH);

            //================================================================= GET ROWS
            //List<org.apache.poi.ss.usermodel.Row> ROWS = wbH.getRows(mySheet);
            List<com.automationanywhere.botcommand.data.model.table.Row> listRows = new ArrayList<>();
            List<String> HEADERS = new ArrayList<>();

            Integer idxs = 0;
            for (Row rw : mySheet) {
                //System.out.println(idxs++);
                List<Cell> listCol = wbH.getColumns(rw);
                //System.out.println(listCol);
                List<Value> rwValue = new ArrayList<>();

                if (HEADERS.size() == 0) {
                    Integer idx = 0;
                    for (Integer colIdx : colsToreturn) {
                        if (hasHeaders) {
                            if (colIdx > (listCol.size() - 1)) {
                                HEADERS.add("");
                            } else {
                                HEADERS.add(wbH.getCellValue(listCol.get(colIdx),0).toString());
                            }
                        } else {
                            HEADERS.add(idx.toString());
                        }
                        idx++;
                    }
                    if (hasHeaders) continue;
                }
                //System.out.println(HEADERS);

                for (Integer colIdx : colsToreturn) {
                    if (colIdx <= (listCol.size() - 1)) {
                        Cell col = listCol.get(colIdx);
                        rwValue.add(wbH.getCellValue(col,0));
                        //System.out.print(wbH.getCellValue(col) + "\t");
                    } else {
                        rwValue.add(new StringValue(""));
                        //System.out.print("EMPTY\t");
                    }
                }
                listRows.add(new com.automationanywhere.botcommand.data.model.table.Row(rwValue));
                //System.out.println("");
            }

            FindInListSchema fnd = new FindInListSchema(HEADERS);

            //System.out.println(HEADERS);
            Table OUTPUT = new Table(fnd.schemas, listRows);
            uteisTest.printTable(OUTPUT,20);


        }catch(IOException e){
            throw new BotCommandException("Error: " + e.getMessage());
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
    public List<Integer> getNumbersInRange(int start, int end) {
        List<Integer> result = new ArrayList<>();
        for (int i = start; i <= end; i++) {
            result.add(i);
        }
        return result;
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

}