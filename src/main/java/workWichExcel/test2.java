import com.sun.org.apache.xerces.internal.xs.StringList;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;
import java.util.*;

public class test2 {
    public static void main(String[] args) throws IOException {
        Map m = new LinkedHashMap();
        Map mRow = new LinkedHashMap();
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

        String track = "e:/TEST.xls";  // files track
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("nature");
        Row row;

        //row
        int i = 0;
        Set<Map.Entry> mapEntrys = m.entrySet();
        Set<String> keySet = m.keySet();
        for (Map.Entry entry : mapEntrys) {
            List natureList = (List) entry.getValue();

            for (int k = 0; k < natureList.size(); k++) {
                if (sheet.getRow(k) == null) {
                    row = sheet.createRow(k);
                    String textFromList = (String) natureList.get(k);
                    System.out.print(textFromList);
                    Cell cell = row.createCell(i, CellType.STRING);
                    cell.setCellValue(textFromList);
                } else {
                    row = sheet.getRow(k);
                    String textFromList = (String) natureList.get(k);
                    System.out.print(textFromList);
                    Cell cell = row.createCell(i, CellType.STRING);
                    cell.setCellValue(textFromList);
                }
            }

            System.out.println();
            i++;
        }

        // wb.write(new FileOutputStream(track, true));
        ((HSSFWorkbook) wb).write(new File(track));
        wb.close();
    }
}
