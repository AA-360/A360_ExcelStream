import com.automationanywhere.botcommand.data.Value;
import com.automationanywhere.botcommand.data.model.table.Row;
import com.automationanywhere.botcommand.data.model.table.Table;
import com.automationanywhere.botcommand.samples.commands.utils.FindInListSchema;

import javax.swing.*;

public class uteisTest {
    public uteisTest(){}

    public static void printTable(Table tb,Integer tam){
        FindInListSchema fnd = new FindInListSchema(tb.getSchema());
        System.out.println("=".repeat(tam*fnd.shemaNames().size()));
        for(String col : fnd.shemaNames()){
            System.out.print(padLeft(col,tam-1) + "|");
        }
        System.out.println();
        System.out.println("=".repeat(tam*fnd.shemaNames().size()));

        for(Row rw: tb.getRows()){
            for(Value col: rw.getValues()){
                System.out.print(padLeft(col.toString(),tam-1) + "|");
            }
            System.out.println();
        }

        System.out.println("=".repeat(tam*fnd.shemaNames().size()));
    }

    public static String padLeft(String s, int n) {
        String val = s;
        if(s.length() > n){
            val = s.substring(0,n-4) + "...";
        }
        return String.format("%" + n + "s", val);
    }

    public void alert(String text){
        JOptionPane.showMessageDialog(null, text, "InfoBox: Title", JOptionPane.INFORMATION_MESSAGE);
    }

}
