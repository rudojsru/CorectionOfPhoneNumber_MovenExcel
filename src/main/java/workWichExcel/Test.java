package workWichExcel;

import com.sun.org.apache.xerces.internal.xs.StringList;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import workWichTextTxtFiles.TextWriterReader;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Test {
    public static void main(String[] args) throws IOException {
        Map m = new HashMap();
        List l = new ArrayList<String>();
        l.add("Kot ");
        l.add("Pies ");
        l.add("żółw ");
        List l2 = new ArrayList<StringList>();
        l2.add("Wrona ");
        l2.add("Sinica ");

        m.put("Ziwotnije", l);
        m.put("Ptach", l2);
        // stworzylem zwykly Map

        TextWriterReader text = new TextWriterReader();
        String track = "e:/TEST.xls";  // files track
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("processedFiles");
        Row row= null;

        int rows;
        rows = sheet.getPhysicalNumberOfRows();

        //
        int j = 0;  // cell
        for (Object list : m.keySet()) {
            int i = 0;   //row
            List list1 = (List) m.get(list);
            for (int k = 0; k < list1.size(); k++) {

                if( row!=sheet.getRow(i)){

                }
                row = sheet.getRow(i);
                // System.out.println(list1);
                String textFromList = (String) list1.get(k);
                System.out.print(textFromList);
                Cell cell = row.createCell(j, CellType.STRING);

                cell.setCellValue(textFromList);
 //              wb.write(new FileOutputStream(track)); // tu jest problem!!!
//                 // kasuje znaczenia kture byli zrobione w poprzednim ządzie.
                i++;

            }
            j++;
            System.out.println();
        }

          wb.write(new FileOutputStream(track,true));
        wb.close();

    }
}
