package workWichExcel;

import com.sun.org.apache.xerces.internal.xs.StringList;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import workWichTextTxtFiles.TextWriterReader;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Test {
    public static void main(String[] args) throws IOException {
        Map m= new HashMap();
        List l=new ArrayList<String>();
        l.add("Kot ");
        l.add("Pes ");
        l.add("Czerepacha ");
        List l2 = new ArrayList<StringList>();
        l2.add("Woron ");
        l2.add("Sinica ");

        m.put("Ziwotnije",l);
        m.put("Ptici",l2);
   // просто создали Map

        TextWriterReader text = new TextWriterReader();
        String track = "e:/TEST.xls";  // files track
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("processedFiles");
   //
        int j = 0;  // cell
        for (Object list : m.keySet()) {
            int i = 0;   //row
            List list1 = (List) m.get(list);
            for (int k = 0; k < list1.size(); k++) {
                // System.out.println(list1);
                String textFromList = (String) list1.get(k);
                System.out.print(textFromList);
                Cell cell = sheet.createRow(i).createCell(j);
                cell.setCellValue(textFromList);
                wb.write(new FileOutputStream(track)); /// сдесь проблемма!!!
                // удалеяется значение предыдущей ячейки на против в столбе и вписывается новое
                i++;

            } j++;
            System.out.println();
        }

        //  wb.write(new FileOutputStream(track));
        wb.close();


    }
}
